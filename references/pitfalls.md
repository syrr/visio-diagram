[中文文档 (Chinese)](pitfalls.zh-CN.md)

# Pitfalls Checklist

Every item below is a problem actually hit in practice, and one that will **waste a lot of time if you don't read it first**. Ordered by severity.

---

## 1. PowerShell cannot instantiate COM (environment level)

**Symptoms**

```
[SandboxError] COM object instantiation can run arbitrary code ...
```

Or `New-Object -ComObject Visio.Application` fails outright.

**Cause**: the execution policy treats COM instantiation as arbitrary code execution and blocks it.

**Fix**: **always use Python + pywin32 instead**.

```bash
pip install pywin32
```

```python
import win32com.client as win32
app = win32.Dispatch("Visio.Application")
```

Late binding (`Dispatch`) is enough; there is no need for `gencache.EnsureDispatch`.

**Corollary**: probe the environment first; don't assume "Visio isn't installed on this machine". Check:

```bash
ls "C:/Program Files/Microsoft Office/root/Office16/VISIO.EXE"
```

Office usually installs Visio and Word together, and Word can be used to update table-of-contents fields, convert to PDF, and test layouts for real.

---

## 2. `Shape.TextBounds` does not exist (API level, the most insidious)

**Symptoms**: you want to tell whether text overflows its box, so you call `shape.TextBounds`:

```
AttributeError: <win32com.gen_py...> object has no attribute 'TextBounds'
```

Neither early nor late binding can reach it. **This property/method simply does not exist in the Visio type library**.

**The danger**: if you wrap it in `try/except` to "probe" for it, it **fails silently** — the checker always reports "no overflow", leaving you believing the layout is fine.

**Fix**: estimate wrapped-line height yourself using **font metrics**.

```python
from PIL import ImageFont
f = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", px)   # point size -> pixels
wpx = width_mm / 25.4 * 96.0
# Lay out character by character against the box width, accumulating line count
asc, desc = f.getmetrics()
need_mm = lines * (asc + desc) / 96.0 * 25.4
```

This is already wrapped as `visio_lib.text_extent_mm()` and `VDraw.check()`.

**Note**: the font name must match the font file used for measuring (Microsoft YaHei → `msyh.ttc`, SimSun → `simsun.ttc`); otherwise the line-count estimate will be off.

---

## 3. Text overflow can only be caught by self-checking

Visio will **not** raise an error or warning when text exceeds its box — it just lets the text spill outside the box, and you only discover it after exporting.

**Fix**: run a self-check before exporting every diagram, making `check()` a fixed step of the figure-production flow.

```python
d.check()     # prints [text overflow] N places, with "how many lines / how tall it needs to be"
```

**Tuning experience** (the actual values used by this skill):

- The left/right inner padding inside shapes, `MARGIN_MM`, from 1.2 → **0.8**, which significantly reduces false-positive overflow;
- Chinese font sizes: box titles 8.5–9 pt, text inside boxes 7–7.5 pt; any smaller and it blurs in A4 body text;
- Rather **make the box taller** than shrink the font size — shrinking the size makes it unreadable after scaled embedding.

---

## 4. Export at 1:1 scale, not at "canvas size"

If you draw at full A4 width (say 21 cm) and then embed into a document whose body column is only 17.5 cm wide, it gets scaled down to 83%, so 7 pt type ends up effectively 5.8 pt and is **illegible in print**.

**Fix**: **draw at 1:1 against the final print width**.

- First measure the body column width (see `docx_embed.content_column_mm()`);
- Design the figure at 95%–100% of the column width (e.g. a 175 mm column → canvas width 150–175 mm);
- Embed at 100%–105% scaling, so the visual size of the type barely changes.

---

## 5. Ignoring the gutter makes tables/figures run off the page

**Symptoms**: the document page is 21 cm wide with 1 cm margins on each side, so you lay out a table at 19 cm — after rendering to PDF the text runs to 21.19 cm, **overflowing the right edge of the page**.

**Cause**: the template's `w:pgMar` carries `w:gutter="851"` (a 1.5 cm gutter), which **takes effect in Word**. The real start of the body text = left margin + gutter = 2.5 cm, so the usable column width = 21 − 1 − 1 − 1.5 = **17.5 cm**.

**Fix**: compute it dynamically from `sectPr`; don't hard-code it.

```python
from docx_embed import content_column_mm
mm, twips = content_column_mm("方案.docx")     # -> (175.0, 9921)
```

Note: if `w:gutterAtTop` is present, the gutter is at the top and does not reduce the width.

---

## 6. Hanging punctuation at the end of a Chinese line gets misread as "overflow"

After rendering to PDF and measuring with text-block boundaries, you will see lines extending about **0.2 cm** past the column width. **This is not a defect** — it is Word's Chinese typesetting behaviour of "allowing punctuation to hang outside the boundary" (punctuation such as `、。；，）` may hang outside the boundary).

**Fix**: allow a 0.3 cm tolerance when judging overflow; or count only lines that end in **non-punctuation**.

---

## 7. Leftover Visio processes

If you forget `app.Quit()` after `doc.Close()`, `VISIO.EXE` stays running in the background. Repeated script runs accumulate multiple instances, eventually causing COM calls to time out or fail.

**Fix**: do both `Close()` and `Quit()` inside `close()`, each wrapped in `try/except` (the document may already be closed).

```python
for fn in (lambda: self.doc.Close(), lambda: self.app.Quit()):
    try:
        fn()
    except Exception:
        pass
```

Also set `app.AlertResponse = 7` to suppress pop-ups — otherwise a single "Save?" dialog can hang an unattended script forever.

---

## 8. `Char.Font` silently fails without quotes

```python
s.CellsU("Char.Font").FormulaU = '"微软雅黑"'   # ✓
s.CellsU("Char.Font").FormulaU = '微软雅黑'      # ✗ no error, but the font does not take effect
```

Strings in Visio formulas must be quoted. Likewise, `CellsU` values are always strings, and booleans use `"0"`/`"1"`.

---

## 9. The absolute physical size in EMF is unreliable — trust only the aspect ratio

**Symptoms**: reading 115.13 mm from the EMF header's `rclFrame`, while a PNG exported from the same figure is 530 px @ 96 DPI = **140 mm**. A 22% discrepancy — which one is right?

**Cause**: when Visio exports EMF, the `rclFrame` it writes implies about **117 DPI**, while PNG uses 96 DPI. The EMF's `rclFrame` (0.01 mm) and `rclBounds` (pixels) do not follow the 96 DPI convention, and different exporters use different values.

**How to verify**: read both `rclBounds` and `rclFrame`; if the implied DPI they yield is not 96, the absolute values are not trustworthy.

```python
info = emf_info("图.emf")
info["implied_dpi"]     # 117.4 → absolute values not trustworthy
info["ratio"]           # 1.0843 → this one is reliable
```

**Fix**: **use only the aspect ratio**; specify the width explicitly from the target column width, and set height = width × ratio.

```python
w_mm = 170.0
h_mm = w_mm * info["ratio"]
```

Measured comparison across four figures — ratios agree to three decimal places while absolute values differ by 22%:

| Figure | EMF `frame_mm` | PNG aspect ratio | EMF aspect ratio |
|---|---|---|---|
| N2 | 113.4 × 116.8 | 1.0306 | 1.0302 |
| N3 | 115.1 × 124.8 | 1.0849 | 1.0843 |
| N5 | 118.4 × 118.4 | 1.0000 | 0.9998 |
| N6 | 115.6 × 116.8 | 1.0132 | 1.0110 |

---

## 10. Word COM throws "Call was rejected by callee"

**Symptoms**:

```
pywintypes.com_error: (-2147418111, 'Call was rejected by callee.', None, None)
```

It appears on seemingly ordinary calls right after `Documents.Add()` — assignments, `Close()`, `SaveAs()`, and the like.

**Cause**: Word is busy (loading in the background, repaginating, or showing a **modal dialog**). COM calls are then rejected outright.

**The danger**: if this exception occurs inside a `finally` block, it will **mask the real outcome of the main flow**; and because `Quit()` never runs, `WINWORD.EXE` is left running in the background — and the leftover instance keeps hogging resources, so every subsequent call fails too, in a vicious cycle.

**Fix**: do both things together.

```python
def _retry(fn, times=6, delay=0.4):
    for i in range(times):
        try:
            return fn()
        except Exception:
            if i == times - 1:
                raise
            time.sleep(delay)

def _quiet(fn):
    try:
        fn()
    except Exception:
        pass

doc = _retry(lambda: app.Documents.Open(path))    # add retries to key calls
...
finally:
    _quiet(lambda: doc.Close(False))              # cleanup never throws
    _quiet(lambda: app.Quit())
```

**Checking for leftover processes** (Windows):

```bash
tasklist //FI "IMAGENAME eq WINWORD.EXE"
taskkill //PID <pid> //F
```

The same applies on the Visio side, but `visio_lib.VDraw.close()` already does `_quiet` handling, and in practice it leaves nothing behind.

**Note**: operating on `Content.Text` immediately after `Documents.Add()` is the easiest place to trigger this — Word is still initializing when a new document is created.

---

## 11. Arrow "drift" — connected to the box, but landing off the edge midpoint

**Symptoms**: the arrow really is connected to the target box, but its landing point is clearly off to one side (often almost touching a corner), making the arrow look as if it has "drifted".

**Real case**: a cross-lane diagram, source box `142–164` (centre 153), target box `100–152` (centre 126).
To cut a corner, the author placed the vertical arrow at `x=147` — the midpoint of the **horizontal overlap of the two boxes (142–152)**,
which looks like "it favours neither side". But the actual rendering shows:

- On the source box it sits at only **23%** of the way across (hugging the left side)
- On the target box it reaches **90%** of the way across (almost touching the right edge)

Both ends deviate from the box centres at once, so visually the arrow "drifts to the right".

**Why it's easy to make this mistake**: taking the midpoint of the overlap zone is reasonable in itself (it guarantees both ends of the arrow stay inside the boxes),
but it ignores the stronger constraint that "visually, the arrow should point at the centre of the box".

**Fix**: when two boxes are offset from each other, **don't use a straight connection**; use an L-shaped polyline, with each end aligned to its box's centre.

```python
# ✗ Straight connection: both ends off-centre
d.v_arrow(147, 38.5, 81.5)

# ✓ L shape: source box bottom-centre → down → left → target box top-centre
d.poly([(153, 38.5), (153, 77), (126, 77), (126, 81.5)], arrow=4)
```

**Self-check**: `scripts/audit_align.py` flags any endpoint that is "connected to one edge alone, yet deviates from that edge's midpoint by > 2mm".
Note that it is a **heuristic**: when several arrows fan out from the same side (each going to a box at a different height), the offset is normal.

**Distinguishing rendering from design**: if you suspect a rendering problem rather than a coordinate problem, overlay the exact geometry from the vsdx onto the exported PNG in red:

```python
# Read PinX/PinY/Width/Height and Begin/End from the vsdx, draw them onto the PNG at the same scale
dr.rectangle([x1*kx, y1*ky, x2*kx, y2*ky], outline=(255, 0, 0), width=1)
```

Red lines coinciding with black lines → the geometry is fine and the design coordinates are the problem; misaligned → it is a coordinate conversion or rendering problem.

**By the way**: a gap of about 0.5mm is often left between an arrow endpoint and the box edge (to keep the arrow from pressing on the box outline),
which is normal practice; but if the gap is so large that the arrow looks like it "doesn't reach the box", shrink it to within 0.3mm.

---

## 12. The bend point of an L-shaped polyline lands on a neighbouring box's edge

**Symptoms**: `audit_align.py` reports that the **start** of an arrowed line segment deviates from some box edge's midpoint by a long way (tens of mm),
yet you clearly wrote it as "aligned to the box centre". Looking at the rendered figure, you find that one stretch of the polyline runs **hugging the bottom edge of the neighbouring box**.

**Real case**: four result boxes side by side; `未命中` ("no hit", x 122–166) needs to connect to `按配置未输出` ("not output per configuration", x 94–134).
Written as

```python
# ✗ Bend point (114, 60.5) lands exactly on the edge of the neighbouring box "命中·数值差>容差" ("hit · value difference > tolerance", x 64–116, bottom edge y=60)
d.poly([(138, 60.5), (114, 60.5), (114, 69.5)], arrow=4)
```

`x=114` falls inside the left neighbour's x range `64–116`, while `y=60.5` is almost exactly its bottom edge `y=60`.
So the horizontal stretch of the polyline runs hugging the left neighbour's bottom edge (visually looking "stuck" to someone else's box),
and at the same time the audit judges the start `(114, 60.5)` as "connected to the left neighbour's bottom edge, yet 24mm off its midpoint".

**Root cause**: thinking only about "the bend point should land in the gap between the two boxes" and forgetting to check **whether the bend point lands on another box's edge**.
This is especially easy to hit when the `y` of a horizontal polyline stretch is chosen too close to the bottom edge of some row of boxes (off by 0.5mm).

**Fix**: the "routing lane" the bend point sits on must keep a **clear gap (≥3mm)** from the rows of boxes above and below,
and the start/end must each align to the source and target boxes' edge midpoints:

```python
# ✓ Route on the y=65 lane: 5mm from the upper row's bottom edge (60), 5mm from the lower row's top edge (70); both ends aligned to box centres
d.poly([(144, 60.5), (144, 65), (114, 65), (114, 69.5)], arrow=4)
```

**Self-check**: `audit_align.py` exposes this kind of problem as "start deviates a long way" —
**when the deviation is especially large (>20mm), suspect first that the bend point is pressing on a neighbouring box's edge**, not that the coordinate conversion is wrong.
Try to make a batch of L-shaped polylines **share the same routing lane** (the same `y`); it looks tidier and is easier to check in one pass.

---

## 13. ★ After switching to OLE, the figure gets squashed (the trap of writing only `Width`)

**Symptoms**: the proportions are fine when the EMF is inserted directly; but once it is "swapped in place for a Visio OLE object" (route C),
the figure **goes flat** — rectangles are stretched horizontally / squashed vertically, visible at a glance. In practice 6 of 11 figures were distorted,
the worst by **+20%** (e.g. N4: EMF aspect ratio 0.973, becoming 1.172 in the document).

**Root cause**: when `InlineShapes.AddOLEObject()` inserts, it sizes by the **OLE object (the Visio document) itself's intrinsic aspect ratio**.
If the script only writes `new.Width = w`, Word **recomputes the height itself** — using Visio's ratio,
not the EMF's. So "the width is right, the height is wrong".

**Why it's hard to spot**: page count, table count, banned words, `w:object` count, `Content_Types`, and orphan headings are **all normal**;
`docx_audit.py` doesn't check this either. You have to compare figure sizes specifically to find it.

**Fix**: before replacing, note down the original figure's `Width` **and** `Height`, write **both** back after replacing, and turn off aspect-ratio locking first
(otherwise the dimension written second will drag the other one along):

```python
rng, w, h = sh.Range, sh.Width, sh.Height     # ← original size comes from the EMF aspect ratio, the only trustworthy one
sh.Delete()
new = rng.InlineShapes.AddOLEObject(FileName=abs_vsdx)
new.LockAspectRatio = 0                       # msoFalse, must be turned off first
new.Width = w
new.Height = h                                # ★ key: explicitly pin the height
```

**Verification (either method works)**:

```python
# Method 1: check the v:shape style in the XML, compare against the EMF aspect ratio; deviation should be ≤0.1%
#   <v:shape id="_x0000_i1026" type="#_x0000_t75" style="width:481.5pt;height:495.5pt">
re.findall(r'style="width:([0-9.]+)pt;height:([0-9.]+)pt"', document_xml)

# Method 2: after exporting to PDF, measure the bounding box of the drawing rectangles (pymupdf)
rs = [d['rect'] for d in page.get_drawings()]      # take those in the figure's y range
aspect = (max(r.x1 for r in rs) - min(r.x0 for r in rs)) / \
         (max(r.y1 for r in rs) - min(r.y0 for r in rs))
```

> ⚠️ When measuring the PDF, **don't use a fixed clip box**: if the figure is taller than the clip, the lower half gets cut off, the computed aspect ratio comes out too large,
> and you are misled into thinking it still isn't fixed. Either use the bounding box of the rectangles from `get_drawings()`, or first confirm the clip covers the whole figure.

**By the way**: the `rclFrame` aspect ratio read by `docx_embed.emf_size_mm()` is reliable (see §9),
and the width/height computed from it during the `合成` ("compose") / `插入` ("insert") stages are **correct**; the error is only in this "swap to OLE" step.

---

## 14. A leftover `VISIO.EXE` makes OLE embed **stale content** (real pitfall, 2026-10-10)

**Symptoms**: after re-running the figure scripts and re-running compose/finalize, one figure in the docx is **still the old content** (old text/old layout),
while the `图表素材/<figure>.vsdx` file on disk is **clearly newly drawn**; the embedded EMF in the "composed but not finalized" version is also new
→ so the problem is only in this **"swap to OLE" step**.

**Root cause**: a headless `VISIO.EXE` process was left over from the previous session. When Word executes
`InlineShapes.AddOLEObject(FileName=vsdx)`, **Visio's OLE server reuses that already-running instance**,
which still has the **old version of the document** open (same path) → so the old content gets embedded into the docx.

**Fix**:
```bash
tasklist //FI "IMAGENAME eq VISIO.EXE"          # first check whether any are left over
taskkill //PID <pid> //F                        # kill it (clear Word too)
# Then re-run: compose → finalize
```

**Prevention / re-verification** (check each figure after embedding; don't just check that "11 figures" is enough):
```python
import zipfile, io
z = zipfile.ZipFile(docx)
# ① OLE preview: word/media/*.emf (vector, text is UTF-16LE)
for n in z.namelist():
    if n.startswith('word/media/'):
        t = z.read(n).decode('utf-16-le', 'ignore')
        assert '旧词' not in t, n            # old word present = the preview is old
# ② Embedded source file: word/embeddings/*.vsdx is itself a zip; read its page XML
for n in z.namelist():
    if n.startswith('word/embeddings/'):
        iz = zipfile.ZipFile(io.BytesIO(z.read(n)))
        t = ''.join(iz.read(p).decode('utf-8', 'ignore')
                    for p in iz.namelist()
                    if p.startswith('visio/pages/page') and p.endswith('.xml'))
        assert '旧词' not in t, n            # old word present = the embedded vsdx is old
```

> In one sentence: **before touching Word/Visio COM, first confirm there are no leftover `VISIO.EXE` / `WINWORD.EXE` processes.

---

## 15. ★ The Y of a connection-point cell points "up" — `add_cp`'s fy is flipped vertically (real pitfall, 2026-10-12)

**Symptoms**: a vertical line starts from the "top edge of the upper box" but **runs straight through both boxes** before stopping; COM prints
`BeginX/BeginY/EndX/EndY` (`Result("mm")`) for each line that completely disagrees with the design coordinates. One incident affected
49 connector lines across 12 figures, yet the naked eye could only see "the line passes through the box" in the rendered figure.

**Root cause**: `Connections.Y{n}` is a **local coordinate** — origin at the shape's **bottom-left corner**, **Y axis pointing up**.
The drawing library's outward convention is "top-left origin, Y pointing down", so writing `=Height*fy` puts `fy=1.0` on the **top edge**
(the opposite of intuition). Nothing in the type library looks wrong about this formula, and Visio raises no error.

**Fix**: inside the library, rewrite it as `=Height*(1-fy)`, and change the meaning of `fy` to "**measured from the top edge**"
(0 = top, 1 = bottom, 0.5 = middle); `side_cps`' direction vectors are synchronized (top → DirY=+1).
Callers need **zero changes**.

```python
self._f(s, "Connections.X%d" % n, "=Width*%.10g" % float(fx))      # fx measured from the left, direction naturally consistent
self._f(s, "Connections.Y%d" % n, "=Height*(1-%.10g)" % float(fy))  # fy measured from the top, must be flipped
```

**Self-check** (more reliable than looking at the figure): after drawing the connectors, use COM to read each line's
`BeginX/BeginY/EndX/EndY` as `Result("mm")` and check them one by one against the design coordinates.
Read the page height from `PageSheet.CellsU("PageHeight").Result("mm")`; **don't guess it**.

---

## 16. ★ Polyline geometry must be translated into `[0,Width]×[0,Height]` — otherwise the selection box detaches from the polyline (real pitfall, 2026-10-12)

**Symptoms**: in Visio, when you select a polyline connector, the **selection box (bounding box) is not on the polyline**,
but appears elsewhere as a transparent empty rectangle; dragging behaves unpredictably too.

**Root cause**: the geometry rows `Geometry1.X{n}/Y{n}` likewise live in local coordinates with "origin at the bottom-left, Y pointing up".
When a polyline goes down/left, the offsets are negative; if you write the negative offsets straight into the geometry rows while still writing
`Width/Height` as a "span" like `max(xs)-min(xs)`, **the shape's width/height become separated from the geometry's actual extent** —
Visio computes the selection box from Width/Height, so it ends up on the other side of the pin.

**Fix**: translate the geometry **as a whole to the bottom-left corner** (the page geometry position and the glued endpoints are completely unchanged),
and set `Pin`/`LocPin` per the **standard Visio 1-D transform** — i.e. `Pin` and `LocPin` **both land at the bounding-box centre**:

```python
xs = [p[0] for p in loc]; ys = [p[1] for p in loc]
dx0, dy0 = min(xs), min(ys)                    # bottom-left corner of the geometry
w, h = max(xs) - dx0, max(ys) - dy0
self._f(conn, "LocPinX", "%.10g mm" % (w / 2.0))            # half width
self._f(conn, "LocPinY", "%.10g mm" % (h / 2.0))            # half height
self._f(conn, "PinX", "BeginX + %.10g mm" % (dx0 + w / 2.0))  # bounding-box centre (anchored on Begin)
self._f(conn, "PinY", "BeginY + %.10g mm" % (dy0 + h / 2.0))
self._f(conn, "Width",  "%g mm" % w)
self._f(conn, "Height", "%g mm" % h)
# Geometry rows write (lx - dx0, ly - dy0)
```

★★ **Don't** write it as `PinX = "=BeginX"` + `LocPinX = -dx0` (LocPin landing on a **corner** of the bounding box) —
that separates the shape's bounding box from the geometry extent. **But note: that is not the real cause of the rectangular outline**;
the real cause is that `Misc.ObjType` was not set to `2` (overturned by a single-variable A/B test on 2026-10-13).
See **§21** (real pitfall, 2026-10-13).

**XML-level acceptance** (does not depend on the Visio GUI; a vsdx is itself a zip):

```python
# Read visio/pages/page*.xml with ElementTree; ★ note: XML <Row IX="n"> starts at 1,
# and the CellsSRC section index does not include the component row — reading with CellsSRC shifts everything by one row, so use the XML.
# In <Cell N="X1" V="0.31496"/>, V is inches; ×25.4 gives mm.
```

Criteria (tolerance 0.02 mm): ① all geometry rows fall inside `[0,Width]×[0,Height]`; ② the four sides of the outer box coincide with the four sides of the geometry;
③ `LocPinX/Y` fall inside the box; ④ the first/last bend points, converted back to page coordinates, equal `BeginX/Y` and `EndX/Y`; ⑤ no diagonal segments (when fully orthogonal).

**Side benefit**: once the bounding box hugs the geometry, the clipping boundary of `Page.Export` also shrinks (no longer inflated by the wrong outer box) —
so changes in PNG/EMF byte counts after re-running the figure scripts are **expected**, not content changes.

---

## 17. Recipe for a single polyline connector (replacing "several combined lines + unglued points")

In Visio, splicing a connector out of several independent lines whose endpoints are not glued to connection points means the line won't follow when you later move a box.
Change to **a single polyline connector + glued to connection points**:

```python
# 1) Create a 1-D shape and glue it to the connection points at both ends
conn = d.connect(src_shape, src_cp, dst_shape, dst_cp)      # already wrapped by the library
# 2) Internally overwrite the coordinate frame, pinning the pin to the Begin point
#    PinX="=BeginX"  PinY="=BeginY"  Angle=0  FlipX=FlipY=0
# 3) Geometry rows: section index 10; row 1 is MoveTo, the rest are LineTo
conn.AddRow(10, -2, 139)        # 139 = visTagLineTo; RowCount(10) includes the component row
```

**Two dead ends (rejected by testing)**: `Page.Layout()` moves boxes that are already placed (`reroute()` can just be emptied;
don't re-lay out the whole thing); under the `pg.Drop(ConnectorToolDataObject)` dynamic-connector procedure, calling `AddRow` afterwards hangs.**

---

## 18. ★★ Merging text into the host shape (`merge_texts`): turning "rectangle + standalone text box" into "rectangle with its own text" (2026-10-12)

Motivation: figures contain many "rectangle + one text box" combinations, and the box itself can hold text; connectors likewise carry their own `Text`.
After merging, the shape count can drop by ~47% (measured across 12 figures: 584 shapes, 302 pure text boxes, of which 274 were inside rectangles).

`VDraw` defaults to `default_auto_merge = True`, automatically calling `merge_texts()` before `save()`/`export()`; it can be turned off with the environment variable `VISIO_MERGE=0` or the class attribute `VDraw.default_auto_merge = False`
(**for producing "unmerged" comparison figures for pixel-by-pixel regression**). `VISIO_MERGE_DEBUG=1` prints where each text box went (which host it merged into / no host / line-label hit and distance).

### 18.1 Writing rich text: **don't use `Characters.CharProps`**

In `Shape.Characters.CharProps(prop, newValue)`, `newValue` is declared as an **integer** in the type library →
**8.5pt gets truncated to 8pt, and colour indexes can't be written at all** (same for early/late binding). The correct approach is to **write the Character section cells directly**:

```python
# Section index 3 = Character; column indexes: 0=Font / 1=Color / 7=Size
host.CellsSRC(3, row, 7).FormulaU = "8.5 pt"     # font size
host.CellsSRC(3, row, 1).FormulaU = "RGB(89,89,89)"  # colour
```

- **Row ↔ paragraph mapping (sentinel method)**: `CharProps(7, 41 + 3*i)` writes a unique sentinel font size for paragraph i,
  then traverse `CellsSRC(3, r, 7).FormulaU` to work the row numbers back out. **Don't** assume "paragraph i = row i".
- **Space before paragraph**: `Characters.ParaProps(4, pt)` (★ the method name is `ParaProps`, **not** `ParagraphProps`).
  Its scope is **all paragraphs within the range passed in** → you must set `c.Begin = offs[i]; c.End = offs[i] + 1` to take
  **only the first character of that paragraph**, otherwise it changes the space-before of all following paragraphs too.

### 18.2 Layout-restoration constants (calibrated by measurement, Microsoft YaHei)

```
Line height   LH = 0.42911·font size − 0.0514  (mm)   ≈ 1.216×font size, residual ≤0.1mm
First-line ink  top of text block + top margin + 1.02 mm        (_INK)
True block height  n×LH + _BH_EXTRA, _BH_EXTRA = 0.52 mm
Inter-paragraph advance  A ≈ LH + _ADV_BONUS, _ADV_BONUS = 0.15 mm
```

- ★★ **True block height ≠ n×line height**: when Visio does `VerticalAlign=1/2` (centred), the "text block height" it uses is **0.52 mm** more than the accumulated line heights
  (same value measured at 8.5pt and 9.0pt). Without compensating, the whole block **shifts down 0.26 mm** (half the amount).
  `VerticalAlign=0` (top-aligned) doesn't use the block height → unaffected, so **only compensate for non-0 alignment**.
- `_INK = 1.02` appears **in both** "block top → first-line ink" and "inter-paragraph advance solved backwards" → **it cancels out completely**,
  so whatever value it takes makes no difference to the final result (but it must be consistent in both formulas).
- Inter-paragraph advance is **quantized** by Visio (about 0.1pt ≈ 0.035mm) → `_ADV_BONUS` anywhere in 0.10–0.30 makes
  almost no difference (the residual is already at the quantization floor); 0.15 is fine.
- **Disproved by measurement**: "when the block height exceeds the available height, Visio stops centring and sticks to the top" — **there is no such clamping rule**
  (counterexample n8). Don't add that compensation.

### 18.3 ★★ Root-cause trap: PIL font metrics **round the font size to an integer** → wrapped line count is under-counted

`ImageFont.truetype(path, size)` takes `size` as an **integer number of pixels**; 8.5pt at 96dpi is 11.333px
→ **rounded to 11px**, so `getlength()` comes out systematically about **~3%** too small (Pillow 12.3.0 rounds even when passed a float).

Consequence: `text_extent_mm()` counts the 13 characters of "读时刻（时间戳计数寄存器）" ("read time (timestamp counter register)") as 1 line (PIL 143px <
limit 145.13px), while Visio actually wraps it into **2 lines** within 38.4mm → the block height is under-counted by 4.1mm → after centring, the text shifts down 1.8mm.

**Fix**: scale the length back up to the true value by the ratio "true px / rounded px":

```python
def _font_scale(size_pt, font_name="微软雅黑"):
    px = max(6, int(round(size_pt * 96.0 / 72.0)))
    return _font(size_pt, font_name), (size_pt * 96.0 / 72.0) / float(px)
# Only f.getlength(...) * sc is the true width
```

### 18.4 ★★ Merging connector labels: the local origin is **not the Begin point**

Writing `conn.Text = "下发"` ("dispatch") to a 1-D connector works (the text appears at the connector's midpoint), but you have to pin the text **back to its original position**:

- **`TxtPinX/TxtPinY` use shape-local coordinates** (Y pointing up). But §17's `_author_path()` translates the geometry
  to the bottom-left corner and uses `LocPin` to pull the pin back to the Begin point → **the local origin = (PinX−LocPinX, PinY−LocPinY)**,
  **not the Begin point**. Writing `cx − BeginX` is only correct when `LocPinX=0` (lines going right/down).
  → Measured: on figure N9, the "上报" ("report up") return line has `Begin=(94,85.5)`, `LocPinX=18`; converting with `Begin` gives a distance
  from the label of **9.86mm** (over tolerance, so the merge is missed), while converting with the local origin gives the true **5.00mm**.
- **`TxtLocPinX` defaults to `TxtWidth*0`** (the **left edge** of the text block sits on the text pin) → you must explicitly write
  `TxtLocPinX = w/2`, `TxtLocPinY = h/2`, otherwise the text **shifts right by ~2mm**.
- `TxtAngle = 0` keeps the text always horizontal (not rotating with the connector angle).
- Hit criterion: the shortest distance from the label centre to the polyline **≤ tol (6mm)** + a **size guard** (width ≤44mm, height ≤12mm)
  → only "labels" get merged, not section titles/legends/explanatory paragraphs.
- Before matching, read the geometry polylines of all 1-D shapes; **the row count must be read with `s.RowCount(10)`**:
  - `s.RowCount("Geometry1")` (a string) → `invalid literal for int()`;
  - For a row that doesn't exist, `CellsU("Geometry1.X3").Result("mm")` **returns 0.0 instead of raising** →
    "stop when you read None" yields a long string of (0,0) phantom points (180 measured in practice). You must truncate by row count.

### 18.5 Regression method: `VISIO_MERGE=0/1` pixel-by-pixel comparison

Run the same script twice (merge off / merge on) and compare the PNGs pixel by pixel. **Judge by "difference rate" and "per-tile displacement"**,
not by the whole-figure difference rate alone (anti-aliasing inflates the noise floor):

```python
# 1) Difference rate: share of pixels with |a-b| > 60; measured across 12 figures 0.27%~1.57% (noise floor from anti-aliasing)
# 2) Per-tile vertical cross-correlation displacement: tiles with ink ≥60 counted → across 12 figures median 0.00px, p95 1px, max 2px
```

The **hard criterion** for "the layout didn't move" is the per-tile displacement, not the difference rate. Residual differences come from anti-aliasing and
space-before quantization (0.035mm level), and are visually invisible.

---

## 19. ★ For two vertically adjacent boxes, the connector endpoints must be "lower box top → upper box bottom" (real pitfall, 2026-10-13)

### 19.1 Symptoms

In figure 9-1, the connector from "授时源" ("timing source", the lower box) to "NTP 授时设备" ("NTP timing device", the upper box) **runs through the entire NTP box**,
with the arrow stopping **above and outside** the NTP box; the expectation was just a short connector between the two boxes.

### 19.2 Root cause

```python
cp["ntp"] = {"bottom": d.add_cp(b_ntp, 0.5, 1.0, diry=1),    # bottom edge y=73 (already defined)
             "top":    d.add_cp(b_ntp, 0.5, 0.0, diry=-1)}   # top edge y=56
cp["ref"] = {"top": d.add_cp(b_ref, 0.5, 0.0, diry=-1)}       # timing source top edge y=77
...
d.connect(b_ref, cp["ref"]["top"], b_ntp, cp["ntp"]["top"])    # ← wrong: the endpoint used the upper box's top
```

The timing source is **below** the NTP box (timing source y 77–94, NTP y 56–73). With the endpoint written as `b_ntp.top` (y=56),
the line runs straight from y=77 to y=56 — and y∈[56,73] in between is exactly **the inside of the NTP box**, hence "piercing the box".
`cp["ntp"]["bottom"]` (y=73) was already defined at the time but **never used** — it was the one that should have been used.

### 19.3 Fix

```python
d.connect(b_ref, cp["ref"]["top"], b_ntp, cp["ntp"]["bottom"])  # lower box top edge → upper box bottom edge
```

Also delete the now-unused `cp["ntp"]["top"]` definition (to avoid misusing it later).

### 19.4 Criterion: ask two questions before choosing endpoints

1. **Which of the two boxes is on top, and which is below?** Read `step_box`'s `ytop/h` to confirm (in this library y points down, top-left origin).
2. **Which edge does the line leave from, and which edge does it enter?** Vertically adjacent → leave from the lower box's **top**, enter the upper box's **bottom**;
   horizontally adjacent → leave from the left box's **right**, enter the right box's **left**.

★ In `add_cp(s, fx, fy)`, **fy=0 is the top edge / fy=1 is the bottom edge** (see §15); `"top"/"bottom"` are just aliases.
**Adding four-way connection points in bulk with `side_cps()` avoids hand-writing fy wrong altogether**:

```python
cp = d.side_cps(b_ntp, "top bottom")   # → cp["top"] / cp["bottom"]
```

### 19.5 Verification trio

- Re-run the figure script → PNG dimensions **unchanged** from before (755×536, content bounding box unchanged), `[text overflow] none`
- Re-run the pipeline → **figure size deviation `1.408 vs 1.409 / 0.096%`, identical to before, character for character** → embedded figure frame size unaffected
- **Visually re-check that page in the final PDF rendering** (Word renders the OLE into the PDF) — this is the most reliable confirmation

---

## 20. ★★ Connectors' **own text**: `d.connect(..., label=...)` (real pitfall ×3, 2026-10-13)

### 20.0 Why do this

When a connector label is made as "rectangle + standalone text box", **selecting the label draws a big rectangular box** (the selection box of a 2-D shape),
and the label doesn't follow when you drag the connector. A **1-D connector's own text**, by contrast, shows only the two end handles and moves with the line.

Criterion (readable from the vsdx alone):

```python
c = {x.get('N'): x.get('V') for x in shape.findall(NS+'Cell')}
one_d = 'BeginX' in c          # True = 1-D connector (own text); False = 2-D text box
```

### 20.1 Usage

```python
d.connect(b_srv, cp["srv"]["r44"], b_sw, cp["sw"]["l44"],
          label="双 10GE", label_color="gray", label_pos=(60, 49.5))
```

- `label_pos=(x, y)` is the **text centre** in page coordinates (mm, top-left origin, Y pointing down);
  if omitted, the polyline bounding-box centre is used.
- `label_align` defaults to 1 (centred), `label_size` defaults to 7pt, `label_color` defaults to the line colour.

> ★★ **`label_pos` is deprecated — don't use it any more** (decided 2026-10-13, see §23): absolute page coordinates **don't move with the line**,
> so one layout change makes all the labels drift off their lines together. Use **`label_at` / `label_gap` / `label_side`** for along-the-line relative positioning instead:
>
> ```python
> d.connect(b_srv, cp["srv"]["r44"], b_sw, cp["sw"]["l44"],
>           label="双 10GE", label_color="gray",
>           label_at=0.5, label_gap=2.2, label_side="up")
> ```

### 20.2 Pitfall 1: writing `Control.TextPosition` makes the text **jump and rotate with the line**

Visio's native "dynamic connector" does carry a named row `Control.TextPosition`, but when you write it via COM in **shape-local coordinates**,
Visio interprets it as **page coordinates**:

```python
# ✗ Don't write it like this
r = c.AddNamedRow(9, "TextPosition", 0)
c.CellsSRC(9, r, 0).FormulaU = "%g mm" % tx     # tx is a local coordinate → interpreted as a page coordinate
c.CellsSRC(9, r, 1).FormulaU = "%g mm" % ty
```

Symptoms: the label jumps somewhere else, and its direction rotates with the connector's direction. **Just set `TxtPin/TxtLocPin/TxtAngle`**.

### 20.3 Pitfall 2: short labels get squeezed into "one character per line"

A connector's text block **carries left/right page margins by default** (about 1.27mm per side) → the actually usable width is ~2.5mm less than `TxtWidth`.
Measured: "授时" ("timing") with `TxtWidth=7.34mm` (text actually needs 4.94mm) still **wrapped character by character**; "双 10GE" ("dual 10GE") with `TxtWidth=11.73mm`
(actually needs 8.97mm) stayed on one line — the 2.5mm difference is the page margins.

```python
self._f(c, "LeftMargin", "0 mm")      # ★ must be zeroed
self._f(c, "RightMargin", "0 mm")
self._f(c, "TxtWidth", "%g mm" % tw)  # also leave 4.5mm of slack (the original 2.4mm wasn't enough)
```

### 20.4 Pitfall 3: **don't** loosen the `_merge_line_labels` heuristic "to rescue a few more labels"

At one point `max_w` was loosened from 44 to 60, the distance criterion was changed to "closest point on the polyline inside the box", and groups with "inconsistent constraints"
were allowed to fall back to connector-label candidates. The result: `best.Text = t` is an **overwriting** write — **when several text boxes merge onto the same connector,
the later ones overwrite the earlier ones**, and the whole text "输出文件复制回本机（回导）" ("copy the output file back to this machine (re-import)") in figure 8-1 was **lost**.

★ The correct approach: **to give a connector its own text, pass `label=` explicitly** (§20.1); don't rely on heuristic guessing.
Keep the heuristic as it was (`max_w=44.0`, centre-distance criterion, no fallback).

### 20.5 Landing checklist

1. Run the figure script → output `[text overflow] none`.
2. Read the vsdx: every label's shape has `1D=True`.
3. Visually check the PNG: **single line, horizontal**, placed consistently with `label_pos`.
   ★ Don't use "PNG width / page width" as the scale when measuring a figure — `Page.Export` crops by the **content bounding box**,
   and the bounding box uses `Pin ± Width/2` (a polyline shape's LocPin is often not at the centre) → the scale comes out wrong.
   The safe approach: work the px/mm back out from the pixel span of a known shape.
4. Re-run the full pipeline → `verify_final` all green (figure size deviation ≤0.1%).
5. **Regress the other 4 figures that use the same library** (byte-for-byte PNG comparison; byte-for-byte `page1.xml` comparison;
   for EMF, only look at the `rclFrame` **aspect ratio** — absolute dimensions are affected by the session/device and can't be compared).

---

## 21. ★★★ A **rectangular outline** appears when a polyline is selected — the root cause is `ObjType` (located and fixed 2026-10-13)

**Symptoms**: when you select a polyline connector we generated, a **rectangular outline is drawn around its bounding box**;
whereas a polyline drawn by hand in Visio (a native Dynamic connector) shows **only endpoint dots + small bend-point squares** when selected,
with no rectangle at all. The user reported explicitly: "the rectangular box still appears when the polyline is selected; it doesn't match the hand-drawn polyline effect".

**★ Root cause = `Misc.ObjType`** (single-variable A/B elimination, 9 control groups):
With the default `ObjType` of `0` (Visio decides for itself) → the shape is rendered as an ordinary 2-D shape, and selecting it draws the bounding box.
A native Dynamic connector has **`ObjType = 2`** (`visLOFlagsRoutable`,
"routable, must be a 1-D shape", see the MS documentation *ObjType Cell, Miscellaneous Section*)
→ selecting it shows only endpoint dots + bend-point squares.

| Variant (identical geometry) | When selected |
|---|---|
| V0 baseline (`ObjType=0`) | **Rectangular outline** ❌ |
| V1 **only** adding `ObjType=2` | **No rectangle** ✅ |
| V2/V3/V4/V5/V7 (`+LockCalcWH` / `+LockHeight` / `+ConFixedCode` / `+LayerMember` / `+TxtPin`, still `ObjType=0`) | All have a rectangle ❌ |
| V8 all added (including `ObjType=2`) | No rectangle ✅ |
| REF Visio native Dynamic connector (`ObjType=2`) | No rectangle ✅ |

**⚠ Unrelated to ShapeTransform (the earlier conclusion was wrong and has been overturned)**:
Both `LocPin=bounding-box centre` and `LocPin=bounding-box corner` **produce a rectangle**
(verified by zooming in on `th_T1_old_zoom.png` / `th_T2_new_zoom.png`; both have a box, only the handle count differs).
The `PinX=BeginX + constant` (centre) form is still kept — it matches Visio's native convention and is harmless.

### The correct solution (already in `visio_lib._author_path`): **order-sensitive, copy it exactly**

```python
# ① First write the transform + geometry correctly per "bounding-box centre"
self._write_geom(conn, [(lx - dx0, ly - dy0) for lx, ly in loc])
# ② First lock the routing, then declare it a "routable 1-D shape"
self._f(conn, "ConFixedCode", "3")
self._f(conn, "ObjType", "2")
# ③ Rewrite the geometry against the "true" local origin after Visio takes over
ox = conn.CellsU("PinX").Result("mm") - conn.CellsU("LocPinX").Result("mm")
oy = conn.CellsU("PinY").Result("mm") - conn.CellsU("LocPinY").Result("mm")
self._write_geom(conn, [(bx + lx - ox, by + ly - oy) for lx, ly in loc])
```

**Three side effects you must know**:

1. **Setting `ObjType=2` triggers Visio's routing engine to take over immediately**: it rewrites the ShapeTransform to the native convention
   `PinX=GUARD((BeginX+EndX)/2)`, `LocPinX=GUARD(Width*0.5)`,
   `Width=GUARD(EndX-BeginX)`, `Height=GUARD(EndY-BeginY)`, and **discards the hand-written geometry**
   (measured: 5 points replaced by 3-point automatic routing) → you must rewrite the geometry **after** it.
2. **`ConFixedCode` must be written before `ObjType`**. Reversed, Visio will **re-route again** during the subsequent recalculation:
   measured on the same horizontal line, `ObjType→write geometry→ConFixedCode` gives 6-point automatic routing;
   `ConFixedCode→ObjType→write geometry` gives the exact 4 points.
3. **Visio clamps any edge shorter than 0.25 in to ±0.25 in** (horizontal lines' Height, vertical lines' Width, both directions for zero-length lines)
   → the local origin `(PinX−LocPinX, PinY−LocPinY)` **≠ the Begin point** (measured: a horizontal line off by 3.175 mm).
   So you **must read the origin back**; don't derive it from Begin.

### Short shapes / free ends get **broken after saving** → self-healing fallback

For **very short** connectors (`ObjType=2` on figure N10's time tick marks at 2.7 mm, arrows at 5 mm) and lines with
**one free end** (figure N8's leader lines), Visio's 1-D minimum-size clamping breaks them
(tick marks measured as shrunk to 0.475 mm; figure N8 gained a diagonal segment).
★ This breakage happens **after `SaveAs`**, while a read-back right after writing looks normal → **on-the-spot validation can't catch it**.

The two lines of defence in `visio_lib.connect()`:

```python
free_end = isinstance(ia, tuple) or isinstance(ib, tuple)
if free_end:
    self._author_path(c, via, objtype=False)      # ① free end: don't add ObjType at all
elif not self._author_path(c, via, objtype=True):
    c.Delete(); c = build()                        # ② on-the-spot validation fails → delete and rebuild
    self._author_path(c, via, objtype=False)
```

Cost: these lines still show a bounding box when selected (legend lines / ticks / arrows / bus leaders).
**Graphical correctness comes first** — rendering must be byte-for-byte unchanged.

### Acceptance points (measured on this project, 2026-10-13)

- Re-running 12 figures: **10 PNGs byte-for-byte identical**, the other 2 with **0 differing pixels**;
  EMF `rclFrame` aspect-ratio deviation across 12 figures **0.000%**.
- `ObjType=2` on **79** connectors in total (= all connectors glued at both ends);
  `ObjType=0` on 52 lines (= all free-end lines). Self-healing fallback triggered **0 times**.
- In the deliverable, the embedded `word/embeddings/*.vsdx` files also contain `ObjType=2` in **79 places**
  → confirming the OLE content was refreshed along with the source, not an old embedded figure.

### How to capture the "selected state" for verification

Use `PrintWindow(hwnd, hdc, 2)` (`PW_RENDERFULLCONTENT`) — it can render an **occluded** window,
which is more reliable than `ImageGrab`; combine it with `win.CenterViewOnShape(shape, 0)` to centre the target,
then crop by the **client-area centre** (`GetClientRect` + `ClientToScreen`; `PrintWindow` output is in window coordinates).
★ `IVWindow` has no `CenterViewOn`, only `CenterViewOnShape(Shape, Flags)`.

---

## 22. ★ The legend `legend()` ignores `w` → the whole row runs off the canvas, and a large blank area appears on the right after export (real pitfall, 2026-10-13)

**Symptoms**: the "main drawing area" in the figure looks narrow, with a large blank area on the right; the legend entries are spaced unevenly, with the last one squeezed far away.
The user's own words: "the main area's width is smaller than the legend's, leaving a lot of blank space on the right".

**Root cause**: `legend()`'s `w` parameter was **never used** — each entry's text box was **hard-coded to 46mm**:

```python
self.text(cx + lw_ + 1.5, y - 2.6, 46, 5.2, label, ...)   # ← 46 hard-coded
cx += lw_ + 1.5 + 46 + gap                                # each entry +79.5mm
```

3 entries starting at `x=6` → the end lands at ≈ `6 + 3×79.5 = 238.5mm`, **far beyond the 170mm canvas**.
And `Page.Export` crops **by content bounds** (not by page) → the overflowing legend is cropped into the figure too,
so the main drawing area (6..166) occupies only about the left 2/3 of the image, with the right 1/3 being "legend text + blank space".

**Fix**: make `legend()` actually distribute evenly by `w` (the whole row runs from `x` to `x+w`):

```python
n = len(items)
cell = (w - gap * (n - 1)) / n          # width of each cell
tw   = max(6.0, cell - lw_ - 1.5)       # text box width (no longer hard-coded 46)
# Each entry: line sample cx..cx+lw_, text at cx+lw_+1.5, then cx += cell + gap
```

When `w <= 0` is passed, it falls back to the old fixed sequential layout (line sample 26 + text 46 + gap).

**Backward compatibility (important)**: `图8_1` ("Figure 8_1") previously passed `w=88` — less than its natural width of 153,
but the old version ignored `w`, so it didn't matter. With the new semantics, `w` must be changed to its **true occupied width of 153**,
otherwise the legend shrinks to 6..94 and the text wraps. After the change, the **PNG and `page1.xml` sha1 were both byte-for-byte unchanged**.
`图9_1` ("Figure 9_1") now passes `w=160` (same width as the content area 6..166).

**How to investigate**: parse the vsdx, compute each shape's `PinX ± Width/2` (the unit is **inches**; ×25.4 gives mm),
and see whether the rightmost shape exceeds "canvas width − margin". If it does, it's most likely the legend / a long text box overflowing,
not the main drawing area really being narrow — don't be fooled by the "crop by content" export.

## 23. ★★ Connector labels "far away from the line" — a general disease: absolute coordinates + bounding-box-centre fallback (real pitfall, 2026-10-13)

**Symptoms**: the user's own words: "in 9.1, the text on the connector is too far from the connector — **this can be a general problem**".
It wasn't a one-off in figure 9-1; both general paths were broken.

**Root causes (three)**:

1. **`label_pos=(x, y)` is an absolute page coordinate that doesn't move with the line.** One layout change (column width / spacing)
   makes all the labels drift off their lines together — figure 9-1 measured **11–14mm** off, even pressing onto neighbouring boxes and crossing the dashed lane borders.
2. **The fallback for `label_pos=None` takes the polyline's "bounding-box centre".** For an elbow-shaped line, the bounding-box centre often
   **isn't on the line at all** (the centre of a 4-segment elbow often falls in an empty triangular area).
3. **The merge path** (`_merge_line_labels`) pins the label at **the position the author placed it**, preserving the offset as-is:
   `VISIO_MERGE_DEBUG=1` measured fig_n8 "生成/供" ("generate/supply") at **5.0/5.5mm**, three labels in fig_n9 at **5.0mm**,
   and fig_n10 "周期开始/时间 →" ("cycle start / time →") at **3.5mm** — they read like standalone notes, not connector labels.

**Fix A — explicit labels in `connect()`: switch to "relative positioning along the line"** (★ recommended; never hand-write absolute coordinates):

```python
d.connect(a, cpa, b, cpb, label="组播")                 # default: midpoint of the longest segment + offset 2.2mm up for horizontal / left for vertical
d.connect(a, cpa, b, cpb, label="组播", label_at=(2, .5))  # midpoint of the 3rd segment
d.connect(a, cpa, b, cpb, label="组播", label_at=0.35)     # at 35% of the whole polyline's arc length
d.connect(a, cpa, b, cpb, label="组播", label_side="down", label_gap=2.6)
```

| Parameter | Meaning |
|---|---|
| `label_at=None` | Automatically take the **longest segment's** midpoint (most stable: longest segment, emptiest on both sides) |
| `label_at=float` | Take a point by **arc-length proportion** along the whole polyline (0=start, 1=end) |
| `label_at=(i, f)` | Take a point at proportion `f` within segment `i` (0-based) — use this to **get around obstacles** |
| `label_gap` | Perpendicular distance in mm from the text **centre** to the segment (default 2.2) |
| `label_side` | `None/'auto'` auto-selects by direction (horizontal segment→**up**, vertical segment→**left**); `'up'/'down'/'left'/'right'/'on'` forces it |
| `label_pos` | Absolute page coordinates, **don't move with the line**; use only when precise positioning is really needed, and visually check the finished product afterwards |

The anchor point sits **on the segment**, then offsets by `gap` along the normal, so the "text↔line" distance stays constant and moves with the connector;
when the layout changes, the labels follow automatically.

**Fix B — merge path: snap to the line**. When merging, **keep the "position along the line" and the "left/right side" unchanged**,
and only converge the "**glyph**-to-line distance" to `LBL_GAP = 0.8mm`:

```python
nd, si, npt, dv = _poly_near(bpts, cx, cy)     # nearest point + direction of the segment it's on
u = (cx-npt)/bd                                # whichever side it was offset to, keep it on that side
reach = a*b / sqrt((b*ux)^2 + (a*uy)^2)        # centre→near-edge distance of the axis-aligned block along u
cx, cy = npt + u * (LBL_GAP + reach)           # near edge of the glyph 0.8mm from the line
```

The glyph box `(gw, gh)` is estimated from the source text box's font size + line count (`_lh` line height + PIL font width measurement),
not from the text box's `w/h` (which includes left/right page margins and would push the label too far away).

**Two supporting lessons (you must visually check the finished product after changing)**:

1. **A vertical segment's `label_gap` must be > half the character width**: a 2-character label's half-width is ≈ 2.47mm, so the default 2.2
   **catches the right edge of the character by 0.3mm** (measured on figure 9-1 "授时" ("timing")); a 4-character label ("授时信号" ("timing signal")) needs 3.0mm or more.
2. **The automatic side can hit a neighbouring block / border**: in figure 9-1, "组播" ("multicast") placed on the left of a vertical segment → it crossed the dashed lane border at x=122,
   and only switching to `label_at=(2, .5)` to land it **above the final horizontal segment** avoided it. When the corridor is narrow, change segment or side;
   don't force it by tuning `label_gap`.

**How to investigate**:
- Merge path: `VISIO_MERGE_DEBUG=1` prints each line label's "original distance → distance after snapping + glyph dimensions".
- Explicit path: parse the vsdx to read the shape's `TxtPinX/TxtPinY`, work the text centre back out with `PinX−LocPinX`,
  and compare its distance against the polyline points from `_line_page_pts()` (more than 3mm off means it has drifted).

**Acceptance (12-figure regression on this project)**: only fig_n8 / fig_n9 / fig_n10 and 图9_1 changed,
with 0.126% / 0.073% differing pixels (concentrated on the labels themselves) and zero change in the other 8 figures;
merged labels at 5.0/5.5/3.5mm all converged to "glyph 0.8mm from the line".


## 24. ★★★ A connector leaves a box from the **wrong side** — it crosses the whole box and is extremely hard to spot with the naked eye (real pitfall, 2026-10-17)

**Symptoms** (the user's own words): "the connector from 『比对（GUI）』 ('Compare (GUI)') to 『报表』 ('Report') **starts in the wrong place, spanning the whole Compare (GUI)**.
This should be investigated as a **common problem**."

**Root cause**: `connect()`'s `ia/ib` only say "glue to which connection point"; they **don't check whether the direction makes sense**:

```python
cp["cmp"] = d.side_cps(box["cmp"], "bottom left")
d.connect(box["cmp"], cp["cmp"]["left"], box["rep"], cp["rep"]["left"])   # ✗
#                     ^^^^^^^^^^^^^^^ target is on the right, yet the line leaves from the left
```

When the two points are at the same height, the connector runs **all the way right from the left midpoint**, crossing the entire source box. The source box has a fill colour → the line is pressed under the box,
with only a short arrow tip showing in the gap on the right, so the finished product looks like "it came across from the left of the box", while **the source script reads perfectly normally**.
Another form: the endpoint is glued to the **left midpoint**, but the path winds into the box and then crosses through to the edge (figure N11's record box in this project).

**Why it must be treated as a common problem**: the side given by `side_cps(...)` in the same script can easily disagree with the actual layout
(who is on the left, who is on the right); one layout change can introduce it, and it **raises no error at all**.

**Fix (two layers, both built in)**:

1. **Block it while writing**: `connect()` now does an **endpoint exit-direction self-check** — when the dot product of the connection point's
   outward normal and the first segment's direction is < 0, it immediately prints:

   ```
   [Connector direction] The Begin end is glued to the left side of shape #18, but the first segment goes the opposite way → the connector will cross that shape.
              Use the side facing the target instead (e.g. connect(a, cp["right"], b, cp["left"])).
   ```
   It only checks **four-way midpoints** (custom connection points have no "facing" semantics, so they're skipped); free points / centre glues are skipped;
   a legitimate route that first goes left and then comes back **won't be falsely flagged** (normal and first segment point the same way).

2. **`d.check()` now checks the whole page as well**: besides text overflow, it also outputs

   ```
   [Connector crossing] Line #7 enters the interior of shape #1 (30×18 @10.0,10.0)
   [Connector self-check] Exit direction 1 place | crossing solid boxes 1 place
   ```
   `audit_links()` uses Liang-Barsky to decide "whether the polyline enters the interior of some **filled** box",
   and excludes two kinds of boxes where "passing through is also normal":
   - **Containers**: area ≥18% of the page area (lanes / system boundary boxes), or containing ≥2 **filled** child boxes;
   - **Unfilled text boxes** (`FillPattern=0`): full-width title boxes; a line passing through the blank space doesn't press on any text.

   ★ The key point in the container test is that "child boxes only count filled ones": a `step_box` still contains 1–2 unfilled text boxes inside it
   (still on the page before `merge_texts`), and if those were counted too, **every step_box would be misjudged as a container**
   and the whole figure would escape checking.

**Self-check method (parse the vsdx, don't start Visio)**:

- A `Cell`'s `V` is in **internal inches** (`U='MM'` is only the display unit) → using it as millimetres is off by a factor of 25.4;
- The connection-point section name in the file is **`Connection` (singular)**, while `Connects` references `Connections.Xn`;
- **`Connections.Y` has its origin at the bottom-left, Y pointing up** → measured from the top, `fy = 1 − Y/Height` (reversed and top/bottom are all wrong);
- `pages.xml` attributes use **single quotes**, so regexes need `['"]`.

**Regression cases (must run after changing the self-check logic)**:

| Case | Expected |
|---|---|
| Left→left (target on the right) | Report "exit direction" + "crossing solid box" |
| Right→left (normal) | None |
| Top→bottom, with a box in between | Report only "crossing solid box" (**should not** report exit direction) |
| Top→bottom, `via` routes around the middle box | None |
| First goes left, then comes back to connect on the right | None (a legitimate route is not falsely flagged) |

**Full results on this project**: 2 places across 15 figures (figure 8-1 "Compare→Report", figure N11 "record box"), 0 places in the other 13;
after fixing, `d.check()` is always "none".

---

## 25. ★★★ A figure looks "squashed / wrong proportions" — do a three-way comparison first, then decide what to change (real pitfall, 2026-10-19, 25 objects)

**Symptoms**: the figure in the document is squashed overall (or stretched), rectangles are clearly not square, and the text in class diagrams is stretched horizontally.
It's easy to judge "the proportions are wrong" by eye, but **changing the wrong thing will wreck the figure completely** (§9 and §13 both proved this: touching the EMF `frame` enlarges and crops the whole thing).

**First, distinguish the three "ratios"**:

| Name | Source | Notes |
|---|---|---|
| Display-box ratio `dr` | In the docx, `<wp:extent cx cy>` (EMU) / `<v:shape style="width:..pt;height:..pt">` | Word lays out the figure by this |
| Bounds ratio `br` | EMF header offset 8, `rclBounds` (4×int32, device units) | The EMF content's true bounding box |
| vsdx ratio `vr` | Union of each Shape's `PinX/PinY/Width/Height` in `visio/pages/page1.xml` (**inches**) | The design truth |

> ★ In a vsdx, a `Cell`'s `V` is always **internal inches**; `U='MM'` is only the display unit (§24 mentioned this too).

**Criteria (a decision tree — don't apply a constant)**:

```
e_b = |br/vr - 1| # deviation of bounds from the design truth
e_d = |dr/vr - 1| # deviation of the display box from the design truth

e_b ≤ 3% → bounds are trustworthy
└ e_d > 3% → only change the display box: H = W / br
e_b > 3% and e_d ≤ 3% → don't touch the frame; re-export the EMF
e_b > 3% and e_d > 3% → still follow "change the docx, not the EMF": H = W / br
```

**Measured on this project** (2026-10-19, 108 figures / 70 OLE objects):

- For 25 objects, `e_d` was consistently ≈ 48% and `e_b` ≈ 0–3% → **all of them were display boxes written wrong**, with `dr = br × 0.52` (1/1.914).
- The root cause of the stable factor: **the EMF `frame`'s X/Y DPI are inconsistent** (measured 48.6 dpi vs 25.4 dpi). The `frame` width is compressed
by a factor of 1.914, so whoever uses the `frame` to compute the display box gets a false ratio of 0.52×.
- ★ **0.52 is not a universal constant**; it only says "this batch of files had their display boxes computed from the frame width". Always use the three-way comparison above;
don't memorize some factor and apply it everywhere.
- After fixing and re-verifying: all 108 figures had `e_d` ≤1.5% (the 12 class diagrams left at 1.7%–3.4% converged after one more refinement pass with `H = W/br`).

**Batch-changing display boxes: all three places must be written; missing one lets Word re-lay out with the old value**:

```python
p = re.sub(r'<wp:extent cx="\d+" cy="\d+"',
f'<wp:extent cx="{int(W*12700)}" cy="{int(H*12700)}"', p) # EMU = 1/12700 pt
p = re.sub(r'<a:ext cx="\d+" cy="\d+"',
f'<a:ext cx="{int(W*12700)}" cy="{int(H*12700)}"', p)
p = re.sub(r'style="width:[\d.]+pt;height:[\d.]+pt"',
f'style="width:{W:.1f}pt;height:{H:.1f}pt"', p)
```

- Old-style VML figures (`<w:object>` + `<v:shape>`): **only style, no extent**;
- DrawingML figures (`<pic:pic>`): only `extent` / `a:ext`;
- OLE objects can appear in either structure; locate the paragraph by rId and change them together.

**The other half of the "too small" problem**: the display-box ratio is right, but the width is only 54% of the column width. Width rule:

```python
W = max(current width, min(design width in mm * 2.83465, column width)) # 2.83465 = pt/mm
H = W / br
if H > 600: # figures taller than the page shrink back by height
H = 600; W = H * br
```

---

## 26. ★★ Word saving normalizes `72.0pt` to `1in` — parsing VML style must support three units (real pitfall, 2026-10-19)

**Symptoms**: the script writes `style="width:340pt;height:72.2pt"`, but after Word opens and saves it once, it becomes
`style="width:340pt;height:1in"`. The audit script parses with `height:([\d.]+)pt` → gets 0 →
the report shows false alarms like "display box 340x**0**" and "aspect ratio −100%", making it look as if the document was damaged.

**Root cause**: when Word serializes VML, it picks the "simplest unit" — a whole 72pt is written as `1in`, and a whole 1in may also be `72pt`;
`mm`/`cm` likewise. This is not corruption; it's normal unit normalization.

**Fix**: always convert units when parsing; don't hard-code `pt`:

```python
def to_pt(v):
m = re.match(r'([\d.]+)(pt|in|mm|cm)?', v)
if not m:
return 0.0
x, u = float(m.group(1)), (m.group(2) or 'pt')
return x * {'pt': 1.0, 'in': 72.0, 'mm': 2.83465, 'cm': 28.3465}[u]

wm = re.search(r'width:([\d.]+(?:pt|in|mm|cm)?)', style)
hm = re.search(r'height:([\d.]+(?:pt|in|mm|cm)?)', style)
dw, dh = to_pt(wm.group(1)), to_pt(hm.group(1))
```

**Related lessons**:
- The same "write 72.2pt → read back 1in" makes **threshold checks flip back and forth**: writing 72.2 and reading back 72.0 is a 0.3% deviation,
which is within normal tolerance; don't rewrite repeatedly "to line up the decimals" — write it once and that's enough.
- Locate a figure's paragraph by **rId → paragraph**, not by the index of "the Nth figure paragraph" (one content change misaligns it).
- After changing, you must **re-export the PDF and re-check** (Word's normalization only takes effect after saving; a static XML check can't see it).

---

## 27. ★★★ A static EMF figure (no vsdx source) has its text clipped — redraw the whole figure; don't expect automatic detection (real pitfall, 2026-10-19)

**Symptoms**: a box title in the figure is missing a character ("规划任务数据处理" ("planning task data processing") shows as "划任务数据处理"),
or the third line of text inside a box presses on the box outline and is cut off halfway. The original figure is an EMF exported by someone else early on, and the docx has no vsdx source.

**Why you can't just "fix it"**: the text in an EMF is **vector outlines / embedded glyphs**, not editable text:

- PyMuPDF **can't extract** spans of this text from the PDF (measured: on the page containing the figure, `get_text()` returns only the body text and captions —
not a single character of the text inside the figure) → the route of "judging in the PDF whether text exceeds its box" is a dead end.
- The PIL font-metrics method (§3) has **extremely unstable thresholds** on this kind of figure: with a line-height factor of 1.30 + full margins it hit 19 places,
and with 1.20 + half margins it hit 0 places — the same batch of figures flips its conclusion → **it can't be used as a criterion**.
- The only option is to **crop each figure and inspect it visually**. Cropping method: take the union bounding box of `page.get_drawings()` over the caption's y range,
crop it, and assemble a contact sheet (4–6 per sheet, DPI ≥240) to look at in one pass.

**Fix = redraw + replace the whole figure**:

1. Redraw from the original figure (normal `visio_lib` flow), making the box wide enough for the **longest text** — don't draw it at a width that merely "looks about right".
After drawing, `d.check()` must reach `[text overflow] none` (it reports "box height/width can't fit" by PIL measurement).
★ But passing `d.check()` ≠ definitely no overflow: it only computes text-box dimensions and can't catch a source-figure design problem like "font size larger than the box",
so when redrawing, **give the box enough height** (in this project, `box()`'s body area = `h - 8.6`; for 2 lines of body text, `h ≥ 17–19mm` is recommended).
2. Reset the display box using the EMF's **bounds ratio** (§25), not the canvas ratio:

```python
b = open(new_emf, 'rb').read()
bd = struct.unpack('<4i', b[8:24])
ratio = (bd[2] - bd[0]) / (bd[3] - bd[1])
H = W / ratio # W keeps the original figure's width (or follows the width rule above)
# A 146x68mm canvas computed by canvas ratio gives a -6.9% deviation ← a real pitfall on this project
```

3. Write back into the docx: with `zipfile`, replace only `word/media/<that figure>.emf` + `word/document.xml`, and copy all other entries **byte-for-byte as they were**
(zero risk to OLE and other EMFs).

**Acceptance**: re-export the PDF, crop the figure and inspect it enlarged (300–400 dpi), confirming ① all text is inside its box ② aspect ratio ≤1.5% ③ centring ≤1pt.

**Results on this project**: 22 "frame≠bounds" objects were suspected of text overflow; after inspecting each one visually, **only 1 really had a problem**,
and the other 21 were fine (class diagrams have blank space on the right by design, with no overflow risk).
→ ★ **A defect list derived from `frame≠bounds` is a hypothesis, not a conclusion; you must inspect visually before reporting numbers.**

---

## 28. ★★ Two acceptance items where the **measurement method itself is flawed**: centring, and table-of-contents page numbers (real pitfall, 2026-10-19)

### 28.1 Checking whether a figure is centred: don't use `Range.Information`

```python
sh.InlineShapes(i).Range.Information(5) # wdHorizontalPositionRelativeToPage
# ✗ All 108 figures return 89.8pt (= the paragraph's left margin), with no distinction at all
```

**The correct approach = the PDF vector-path method**: an EMF is played back into the PDF as vector paths; take the union bounding box of the
**normal-sized** paths in `page.get_drawings()` and compare it with the column centre:

```python
rects = []
for dr in page.get_drawings():
r = dr['rect']
if r.width <= 0.5 and r.height <= 0.5: # exclude points/short lines
continue
if r.width > COLW * 1.06 or r.x0 < col0 - 30 or r.x1 > col1 + 30:
continue # ★ exclude the EMF's oversized clipping box
rects.append(r)

# Slice by caption y-range, taking only paths "after the previous caption, before this caption"
cand = [r for r in rects if r.y0 >= low - 2 and r.y1 <= cap.y0 + 3]
x0, x1 = min(r.x0 for r in cand), max(r.x1 for r in cand)
offset = (x0 + x1) / 2 - (col0 + col1) / 2 # ≤1pt counts as centred
```

**Notes**:
- **Exclude the table-of-contents / list-of-figures pages** (page text containing both `....` and `图 ` ("Figure "));
- For bitmap figures (`<pic:pic>`, not EMF), supplement the measurement with the bbox from `page.get_image_info()`;
- The "ink bounding box sliced by caption band" approach gives **false positives**: text inside boxes in the EMF is also extracted as text lines,
squeezing the band into a narrow strip (on this project it once falsely reported 2 well-centred figures as 67–70pt off to the left).

### 28.2 Verifying list-of-figures page numbers: take the page number from the "footer block", not from text-line order

```python
# ✗ Wrong: assuming the page number is in the last few lines of get_text()
for line in reversed(page.get_text().strip().split('\n')[-6:]):...
# The footer block may sit 2nd in PyMuPDF's block sequence → you pick up a number from the body text, and the whole table is falsely flagged
```

```python
# ✓ Right: take the footer region by coordinates (A4 height 841.9pt)
for b in page.get_text('blocks'):
if b[1] > page.rect.height - 70: # bottom 70pt band
s = b[4].strip()
if s.isdigit():
label = int(s); break
```

**Why verification is necessary**: changing figure sizes changes the page count (on this project, 204 → 201 pages),
so after `doc.Fields.Update()` + `TablesOfContents/Figures.Update()` you must **re-export the PDF and compare entry by entry**
"table-of-contents page number vs actual page number in the body". On this project, all 191 list-of-figures entries ended up 100% consistent.

**The correct way to export a PDF**: `doc.SaveAs2(pdf_path, 17)` (wdFormatPDF) via a **temporary copy**;
`ExportAsFixedFormat` **silently fails** in some environments (no error, no file produced). The path must be `os.path.abspath()`,
otherwise Word resolves the relative path under `C:\Windows\System32\` and reports "couldn't find your file".

## 29. ★★★ Text at the bottom is clipped / the figure shrinks into the top-left corner — the three sets of dimensions in the EMF header must agree with each other (real pitfall, 2026-10-25, 17 COMAC figures)

### 29.1 Symptoms and misdiagnosis

The last line of text in a figure (caption, bottom note, footnote) shows **only its top half** in the Word / docx preview, while the same-source PNG
is completely normal. It's easy to misdiagnose this as "the figure source's layout was clipped" and go and change the Visio source figure — **but the figure source is fine**.

### 29.2 True cause: `Page.Export`'s rclBounds is about 9 px shorter than the actual bitmap

| Quantity | Measured value (figure 11-7) |
|---|---|
| Same-source PNG | 601 × 342 px |
| EMF `rclBounds` | 600 × **333** px ← 9 px short (≈2.4 mm) |

Visio underestimates the lower edge of the text (glyph bounding box vs ink) when computing the clipping box; **the bitmap draws the characters in full,
but the EMF's bounds are 9 px short**. Word/OLE maps by rclBounds → the lower half of that bottom line of text is cut off.

```python
# Criterion: bounds height + 9 ≈ PNG height (widths are basically the same; a 1 px difference is rounding)
rb = struct.unpack_from("<iiii", buf, 8)
print(rb[3] - rb[1], Image.open(png).size[1]) # 333 342
```

**Fix**: after exporting, use the same-source PNG's pixel dimensions to top up rclBounds (`_fix_emf_size(path, W, H, png=...)`,
and `export_all()` already orders PNG **before** EMF). Add one more safeguard on the figure-source side: use
`valign=0` (top alignment) for text at the bottom, so it doesn't sit exactly on the clipping-box edge.

### 29.3 The other half: rclFrame and szlMillimeters must use the same calibration

A renderer has two criteria for an EMF's physical size, and **when they disagree, Word takes `szlDevice ÷ szlMillimeters`**:

| Source | rclBounds ÷ rclFrame | szlDevice ÷ szlMillimeters | Result |
|---|---|---|---|
| Visio's original export | 3.78 px/mm | 14.8 px/mm (2560/173) | **Figure shrinks to 1/4 of the object box, stuck in the top-left corner** |
| Correct (what this library now outputs) | 3.78 | 3.78 | Fills the object box |

So `_fix_emf_size()`, besides writing rclFrame / szlMillimeters, **must also write szlDevice as the content pixel dimensions** (`struct.pack_into("<ii", buf, 72, bw, bh)`), so all three agree.

### 29.4 Re-check list

1. `bounds == PNG dimensions` (±1 px);
2. `bounds_w / frame_w_mm` equals `dev_w / mm_w` (±0.05);
3. In the docx, crop a 500 dpi strip **3 mm above the figure caption** and look at it — half-cut characters can't be seen at 100 dpi;
4. Compare vector extents figure by figure against the previous PDF known to be correct; the ratio should fall in 0.96–1.04.

### 29.5 When replacing a preview image, don't go by the `imageN` number

In a docx re-saved by Word / WPS, the numbering of `word/media/imageN.emf` is **unrelated to the order of appearance in the body**
(the later-inserted figure 11-7 is `image17.emf`, yet appears 10th). Replacing by number swaps the wrong figures —
you must take the `rId` → media file in the order of `w:object` in `document.xml`, then match by **figure caption**.

---

## 30. Curved connectors and arrow size (verified 2026-11-19)

### 30.1 `Page.DrawSpline` **cannot be made to work** under COM

Both early and late binding were tried:

```python
page.DrawSpline((0.5, 0.5, 1.0, 1.2, 1.5, 1.4, 2.0, 1.0), 0)
# pywintypes.com_error: (-2147352562, 'Invalid number of parameters.')
page.DrawSpline(VARIANT(VT_ARRAY|VT_R8, vals), 0)
# TypeError: Objects for SAFEARRAYS must be sequences (of sequences), or a buffer object.
```

No way of passing the parameters works (`xyArray` needs Visio's private SAFEARRAY layout). **Don't burn time on this**:
curves always use `connect(..., via=[control points], smooth=N)` — this library samples a Catmull-Rom spline through the control points
into a dense polyline (`catmull_rom()`), N points per segment, with the geometry rows still ordinary `LineTo`.

Costs and cautions:

- A curve has no draggable control points (in Visio you can only drag it as a whole); to change its shape, change the script's `via`;
- **Don't repeat control points** (when gluing to a connection point, an endpoint coincides with the first/last point) → `catmull_rom()` de-duplicates internally,
  otherwise a zero-length segment distorts the tangent, and the end hooks back on itself and the arrow direction reverses with it;
- The first segment's direction must face the connection point's outward normal, otherwise `_check_end_dirs()` still raises the alarm (curves are not exempt).

### 30.2 A drawn 1-D shape has **no `ArrowSize` cell**

Shapes generated by `DrawLine` / `DrawPolyline` have only `BeginArrow` / `EndArrow` in the Line section,
with **no `ArrowSize`** (`CellsU("ArrowSize")` reports `Unexpected end of file`;
`CellsSRC(1,0,6)` lands on the geometry section). Consequences:

- The arrow is always Visio's "medium" size, and **does not scale with `LineWeight`** (measured: the arrow is the same size at 0.5pt and 0.75pt);
- There are only three ways to make an arrow smaller: ① leave enough length for the arrow (preferred — don't cram an arrow into a 3 mm gap);
  ② make the arrow land in the **tangent direction** of the circle/box (approaching horizontally is cleaner than approaching at an angle); ③ hand-draw a small triangle
  (three `poly` points) as the arrow, with `EndArrow=0`.

★ Judge "whether the arrow is too big" at **96 dpi** (how it actually looks in Word at 100% zoom).
Looking at a small arrow at 300 dpi, magnified 3×, makes any 0.75pt line's arrow look like a blurry blob, and you can easily waste half a day changing things for nothing.

---

## 31. Replacing an **existing** OLE figure in place in a document (rework scenario, verified 2026-11-19)

In a docx delivered by the client (甲方, the client side) / a previous version, figures are often Visio OLE objects wrapped in `<w:object>`. When replacing a figure:

### 31.1 You must replace **two places**; missing one gives "the displayed figure and the figure you get on double-click don't match"

| Slot | Purpose |
|---|---|
| `word/embeddings/*.vsdx` (or `Microsoft_Visio_Drawing.vsdx`) | The source object opened when you double-click to activate it |
| `word/media/*.emf` | **This cached preview is what Word normally draws**, without opening Visio |

Replace only the vsdx → the page still shows the old figure; replace only the EMF → double-clicking still opens the old figure.

### 31.2 The display box width/height must be rewritten from the new EMF

In `<w:object>`, `<v:shape style="width:439.5pt;height:227.5pt">` is the display box. If the figure is redrawn so its aspect ratio changes
but this isn't updated, Word will **stretch** it (common symptom: the figure is squashed, the characters look thin). How to change it:

```python
STYLE_RE = re.compile(r'(style="width:)([\d.]+)(pt;height:)([\d.]+)(pt)')
# Only do one replacement inside <w:object>…</w:object>; don't touch other images
```

Height = width ÷ EMF aspect ratio, with the ratio taken from the EMF header's `rclFrame` (0.01 mm) — **don't use** the old value in the VML
`style`, and don't use `w:dxaOrig` (that's the original object size, unrelated to the display box).

### 31.3 The display box width must not exceed the body column width

Measured on this machine: `w:pgMar` is 1797 twips on each side → column width 8313 twips = **415.7pt**. Old figures are often
439.5pt (8mm past the page margin); when replacing a figure, bring it inside the column width as well, fixing the "figure wider than the table" layout problem at the same time.

### 31.4 Ready-made script

```bash
python scripts/replace_ole.py 方案.docx 图N2.emf 图N2.vsdx      # default display box = column width − 0.5pt
python scripts/replace_ole.py 方案.docx … --width-pt 400        # give it explicitly when a narrower one is needed
```

The script only handles "one figure, one object"; for multiple figures, per §29.5, **match the `rId` in `document.xml`'s `w:object` to the media file, then match by figure caption** — don't guess by the `imageN` number.
