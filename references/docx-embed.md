[中文文档 (Chinese)](docx-embed.zh-CN.md)

# Embedding Visio Figures in Word

Three routes: **A** (insert picture, static vector) → **C** (OLE object, double-click to edit) → **B** (hand-built OOXML, fallback).
Which one to pick depends on "does the client want to double-click the figure in Word to edit it":

| Need | Route | Result |
|---|---|---|
| Just needs the figure visible and sharp when scaled | A / B | `<w:drawing>` + EMF |
| Client wants to **double-click into Visio to edit** | **C** | `<w:object>` + `<o:OLEObject ProgID="Visio.Drawing.15">` |
| No Word environment, pure script composition | B | Same as A |

**Route C is best layered on top of A**: first insert the EMF figures with A/B (layout and widths all set), then use C to swap each figure **in place** for an OLE object — that way the layout logic is fully reused and no sizes need recalculating.

---

## Route C: Convert to Double-Click-Editable Visio Objects (OLE)

A common client requirement: "the figure must open in Visio for editing when double-clicked in Word." In that case you cannot insert an EMF; you must embed an **OLE object**.

**Approach**: Word COM's `InlineShapes.AddOLEObject()`, replacing existing pictures **in place** while keeping their original display width.

```python
import win32com.client
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False; word.DisplayAlerts = 0
doc = word.Documents.Open(DOC, ReadOnly=False)
shapes = list(doc.InlineShapes)                 # Must snapshot first; the collection changes during replacement
assert len(shapes) == len(FIGS), "图片数与预期不符"
for sh, fig in zip(shapes, FIGS):
    rng, w, h = sh.Range, sh.Width, sh.Height   # Remember Range + width + height first
    sh.Delete()                                 # Then delete, otherwise Range becomes invalid
    new = rng.InlineShapes.AddOLEObject(FileName=abs_vsdx)
    new.LockAspectRatio = 0                     # msoFalse: must turn it off first
    new.Width = w                               # Restore width
    new.Height = h                              # ★ Height must also be restored (see pitfall 3)
for toc in doc.TablesOfContents:
    toc.Update()                                # Figure count changed, page numbers must be re-laid-out
doc.Save(); doc.Close(False); word.Quit()
```

**Key points / pitfalls**:

1. **Order-sensitive**: `InlineShapes` is ordered by appearance in the document, and `FIGS` must correspond to it one-to-one. Snapshot with `list()` first, then modify; deleting while iterating skips items.
2. **Get `Range` before `Delete`**: after deleting the shape, `sh.Range` is invalid, so the insertion point must use the `rng` captured before deletion.
3. **★ Both width and height must be restored manually**: `AddOLEObject` inherits neither the original width nor the original height — it sizes by the **OLE object (Visio)'s own intrinsic aspect ratio**. If you only set `Width`, Word recalculates the height from it, the result disagrees with the EMF figure extent, and the **figure is squashed (measured up to +20%)**, while routine checks on page count, tables, and banned words all pass green and catch nothing. So after `LockAspectRatio = 0` you must write back **both** `Width` and `Height` (the original height taken from the figure before replacement, i.e. the EMF ratio). See `pitfalls.md` §13.
4. **What gets embedded is a `.vsdx`**, not an `.emf` — so keep the `.vsdx` source files (routes A/B only need the `.emf`).
5. **The TOC must be `Update()`d**: after pictures become OLE objects the pages reflow, and the table-of-contents page numbers will all be wrong.
6. **After finishing, remove `<w:updateFields/>` from `settings.xml`**, otherwise the client gets a "do you want to update fields" prompt on opening.
7. **Word rewrites the EMF preview** (`word/media/*.emf` byte counts change); this is normal. When verifying, do not compare the embedded EMF's sha1 against the source EMF — compare the **`.vsdx` object structure** instead.

**Verification** (mandatory after the swap):

```python
d = zipfile.ZipFile(docx).read('word/document.xml').decode('utf-8')
d.count('<w:object')          # == figure count
d.count('Visio.Drawing.15')   # == figure count
d.count('<w:drawing')         # == 0 (no bare pictures should remain)
len([n for n in z.namelist() if 'embeddings/' in n])   # == figure count
len([n for n in z.namelist() if 'media/' in n])        # == figure count (EMF previews)

# ★ Figure-extent check: the width/height ratio in v:shape's style must match the EMF aspect ratio (deviation ≤0.1%)
import re
from docx_embed import emf_info
styles = re.findall(r'style="width:([0-9.]+)pt;height:([0-9.]+)pt"', d)
for (wp, hp), emf in zip(styles, EMFS):        # EMFS in the same order as FIGS
    assert abs(float(wp) / float(hp) / emf_info(emf)["ratio"] - 1) < 0.001
```

Then open Word once more and read `doc.InlineShapes.Count` and the count of `Type==1` (wdInlineShapeEmbeddedOLEObject) to confirm all of them are OLE.

**Note**: Word **normalizes XML** on save (e.g. `<w:b />` → `<w:b/>`, attribute order changes). When verifying with string counts, match both spellings; don't write only one.

---

## Route A: Word COM Insertion (Recommended)

Word natively supports EMF with vector fidelity, and you **don't have to touch OOXML**.

**Approach**: leave placeholders in the document, then batch-replace them after drawing the figures.

```markdown
... 系统工作模式见下图。

{{FIG:N2}}

图 9-2 系统工作模式与数据流向图
```

```bash
python scripts/word_insert.py 方案.docx --placeholder "{{FIG:N2}}" --image 图N2.emf
```

The script: finds all `{{FIG:N2}}` → clears the placeholder → inserts the picture at the insertion point → sets the width to the body column width → sets the height from the EMF aspect ratio → centers the paragraph → updates fields → saves as a new file.

Default width = **body column width − 5 mm** (`content_column_mm()` computes it automatically, binding edge already accounted for).

It can also be used as a library:

```python
from word_insert import insert_at_placeholder
insert_at_placeholder("in.docx", "{{FIG:N3}}", "图N3.emf", width_mm=170, out="out.docx")
```

**Advantages**: independent of how the document was generated; EMF sizes are parsed by Word itself; vectors are not lost.
**Prerequisite**: Word installed on this machine.

---

## Route B: Hand-Built OOXML

For scenarios where a docx is composed from scratch by script (no Word, or fully programmatic).

### B1. Read the EMF Figure Extent from the File Header

PIL **cannot read EMF** (it throws an exception), and `python-docx`'s `add_picture` uses PIL internally, so you **cannot insert an EMF directly**. You must get the size from the EMF header yourself:

```python
from docx_embed import emf_info
info = emf_info("图N2.emf")
# {'bounds_px': (523,539), 'frame_mm': (113.4,116.8),
#  'ratio': 1.0302, 'implied_dpi': 117.2, ...}
```

EMF header structure (ENHMETAHEADER):

| Offset | Field | Notes |
|---|---|---|
| 0 | `iType` | Always 1 |
| 4 | `nSize` | Header length |
| 8 | `rclBounds` | 4×int32, **device units (pixels)** |
| 24 | `rclFrame` | 4×int32, **units of 0.01 mm** |
| 40 | `dSignature` | Should be `" EMF"` (`0x464D4520`) |

When writing `wp:extent`, convert to EMU: `cx = mm * 36000`.

> ⚠ **The absolute physical size of an EMF is unreliable**. In practice, EMF exported from Visio has an implied **~117 DPI** from `rclFrame`, while a PNG exported from the same figure is **96 DPI** (physical sizes differ by about 22%). But the **aspect ratios are exactly the same** (difference < 0.1%).
>
> **Correct approach**: take only the aspect ratio, specify the width explicitly from the body column width, and set height = width × `ratio`. Do not trust the absolute value of `rclFrame` directly — otherwise the figure will shrink or grow for no apparent reason. `word_insert.py` and `docx_audit.py` both do it this way.

### B2. Add the Content_Types Declaration

```python
from docx_embed import ensure_content_type
ensure_content_type("方案.docx", "emf")     # Adds <Default Extension="emf" ContentType="image/x-emf"/>
```

### B3. Three OOXML Pitfalls You Will Hit

#### Pitfall 1: `<Default>` Must Be Written Inside `<Types>`

`[Content_Types].xml` has only one root node, `<Types>`. If you insert `<Default>` after the XML declaration but before `<Types>`, the file has two root-level elements — and Word reports **"file is corrupted"**.

```python
m = re.search(r"<Types[^>]*>", ct)
ct = ct[:m.end()] + decl + ct[m.end():]      # ✓ Insert after the <Types> opening tag
```

#### Pitfall 2: `mc:Ignorable` May Only Reference Declared Prefixes

The root element looks like:

```xml
<w:document xmlns:w="..." xmlns:mc="..." mc:Ignorable="w14 wp14">
```

If you parse with `ElementTree` and serialize again, **unreferenced namespace declarations are dropped**, but `mc:Ignorable` still lists `w14` — violating the MCE spec, and Word again reports **"file is corrupted"**.

Fix: after serializing, trim `mc:Ignorable` again, keeping only prefixes still declared.

```python
head = s[:s.find(">", s.find("<" + root_tag)) + 1]
declared = set(re.findall(r'xmlns:([\w\d]+)="', head))
keep = [p for p in ignorable.split() if p in declared]
```

`docx_audit.py` already has this check built in, for both `document.xml` and `styles.xml`.

#### Pitfall 3: The Attribute Name for `w:tblW` / `w:tcW` Is `w:w`, Not `w:val`

```xml
<w:tblW w:w="9921" w:type="dxa"/>     <!-- ✓ -->
<w:tblW w:val="9921" w:type="dxa"/>   <!-- ✗ -->
```

Written as `w:val`, Word reads `PreferredWidth = 0`, and **every column collapses to about 0.4 cm** (one character per line), ballooning the page count from 32 pages to 201.

What makes it confusing is that `w:gridCol` really does use `w:w` (that one is correct), so not every width attribute is written wrong — only `tblW` / `tcW` are easy to write as `w:val` by habit.

`docx_audit.py` already has this check built in.

---

## Pre-Delivery Self-Check

```bash
python scripts/docx_audit.py 方案.docx
```

The static part (no Word needed) checks:

- All XML parts are well-formed
- Whether `<Default>` is inside `<Types>`, and whether the extensions in `word/media` are all declared
- `mc:Ignorable` prefix consistency
- `tblW` / `tcW` attribute names
- Bold run counts
- Body column width (including binding edge)

The Word live-test part reports: page count, paragraph and table counts, **widest table vs. column width**, each figure's size, and bold paragraph counts inside and outside tables.

Decision criteria:
- Widest table ≤ body column width → pass (0.05 cm tolerance)
- Bold inside tables = 0 → pass
- Bold outside tables = heading paragraph count → normal (heading styles are bold by themselves)

---

### Figure Layout Acceptance (Figure Extent / Centering / TOC Page Numbers)

`docx_audit.py` only looks at text and tables; the three figure-specific items below must be run separately. All three must be done **after exporting to PDF**.

**① Figure-extent ratio — three-way comparison, don't guess**

```python
# Compare three things: display box (docx), bounds (EMF header), vsdx union (design ground truth)
# Criteria and decision tree: see pitfalls.md §25
dr = dw / dh                     # <wp:extent> or <v:shape style>; mind in/pt/mm units (§26)
br = (bd[2] - bd[0]) / (bd[3] - bd[1])   # struct.unpack('<4i', emf[8:24])
vr = vsdx_union(path)            # PinX/PinY/Width/Height union, in inches
```

Decision: `|br/vr-1| ≤ 3%` and `|dr/vr-1| > 3%` → the display box is written wrong; change to `H = W / br`. When reporting, list the `dr / br` deviation for the whole table; **anything >1.5% must be looked at** (first round on this project: 25 figures at a constant −48%; after fine-tuning, 12 remaining at 1.7%–3.4%; only counted as done once all converged to <1.5%).

**② Horizontal centering — PDF vector-path method**

```python
rects = [d['rect'] for d in page.get_drawings()
         if d['rect'].width > 0.5 and d['rect'].width <= COLW * 1.06]   # Exclude oversized EMF clipping boxes
# Take the union bounding box over the caption's y-range and compare against the column center; deviation ≤1pt
```

★ Do not use `Range.Information(5)` (it always returns the paragraph's left margin for inline shapes); and do not slice bands by the "ink bounding box" (figure text gets treated as body text, causing false positives). See `pitfalls.md` §28.1 for details.

**③ Table of contents / table of figures page numbers**

Changing figure sizes necessarily changes the page count (on this project, 204 → 201). After updating fields you **must** re-export the PDF and compare item by item: each TOC entry's page number vs. the page where the caption actually appears in the body.

```python
doc.Fields.Update()
for i in range(1, doc.TablesOfContents.Count + 1):   doc.TablesOfContents(i).Update()
for i in range(1, doc.TablesOfFigures.Count + 1):    doc.TablesOfFigures(i).Update()
doc.Repaginate(); doc.Save()
```

Page numbers must be read from the **footer blocks** (pure-number blocks with `y0 > page height − 70` in `blocks`); don't take the last line from `get_text()` line order (see `pitfalls.md` §28.2).

**How to export the PDF**

```python
shutil.copy2(docx, tmp)                        # Work on a temp copy; the original document is untouched
doc = app.Documents.Open(os.path.abspath(tmp)) # ★ Must be abspath, otherwise Word looks in System32
doc.SaveAs2(os.path.abspath(pdf), 17)          # 17 = wdFormatPDF
```

`ExportAsFixedFormat` **fails silently** in some environments (no error, no file produced); prefer `SaveAs2`.
