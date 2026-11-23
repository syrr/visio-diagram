---
name: visio-diagram
version: "1.0.1-beta"
description: Generate natively editable Visio diagrams cross-platform and embed them in Word documents. Without Visio (Linux/macOS/CI/sandbox), use the XML backend to produce editable .vsdx + PNG previews and write them back into docx; on Windows with Visio installed, use the COM route to produce vector .emf. Use when the user asks for architecture diagrams, flowcharts, swimlane diagrams, module relationship diagrams, or data flow diagrams, or asks to embed a diagram in a docx, or mentions Visio, vsdx, technical proposal diagrams, or overall design document illustrations. Suitable for engineering technical documentation diagrams, including CJK, with built-in text-overflow self-check and docx layout verification.
agent_created: true
---

[中文文档 (Chinese)](SKILL.zh-CN.md)

# Visio Diagrams and Document Embedding

## Two routes (pick the right one before you start)

| Scenario | Which route | Entry point |
|---|---|---|
| No Visio (Linux / macOS / CI / agent sandbox) | **XML backend**: pure Python, generates editable vsdx + PNG preview, can write back into docx | `scripts/xml_backend/`, see `references/cross-platform.md` |
| Windows + desktop Visio, need vector EMF or live Visio rendering | **COM route**: drives the local Visio, produces `.vsdx` + vector `.emf` + `.png` | `scripts/visio_lib.py` (the rest of this document) |
| Windows + Visio, want to verify XML backend output | Combined: XML backend generates → `visio_finish.py` verifies/finishes (recalculate, export EMF) | `scripts/xml_backend/visio_finish.py` |

★ Rule of thumb: **does this machine have VISIO.EXE?** If not, take the XML backend (don't try to install COM dependencies); only take COM when Visio is present and you need EMF.
The XML backend's output is likewise natively editable Visio objects (connection points, glue formulas, and connector-carried text all present); the only difference is that the preview is PNG rather than vector EMF.

### XML backend quick start (cross-platform, 30 seconds)

```bash
pip install pillow
# 1) Standalone vsdx + PNG preview (write FIGS following assets/example_xml_flowchart.py)
python scripts/xml_backend/build_vsdx.py assets/example_xml_flowchart.py out/
# 2) Redraw Visio diagrams already embedded in a docx (the original is read-only; no Word/Visio needed)
python scripts/xml_backend/inspect_visio_docx.py 方案.docx              # diagnose current state
python scripts/xml_backend/extract_vsdx.py 方案.docx work/templates     # extract per-diagram templates
python scripts/xml_backend/build_docx.py 方案.docx work/figspecs.py work/templates out.docx
python scripts/xml_backend/inspect_visio_docx.py out.docx               # re-check
# 3) Tests (no Visio needed)
python -m unittest discover -s tests -v
```

Fonts default by operating system (Windows=Arial+Microsoft YaHei / macOS=Helvetica Neue+PingFang / Linux=fontconfig detection); the `VISIO_FONT_LATIN/CJK` and `VISIO_FONT_*_FILE` environment variables can override them. XML backend themes (default `tech-cn`; leave it unset to keep it): a single diagram with `Fig(W, H, theme="mono")`, a whole batch with `build_vsdx.py / build_docx.py --theme mono`. `mono` is also available (black-and-white print); use one theme per document — a theme only changes colors, never meaning. See **`references/cross-platform.md`** for the full description and known limitations. Everything below is COM-route-only content.

Both routes are for engineering document diagrams that need **precisely controlled layout**. They are not for quick throwaway sketches (inline SVG is faster for that).

## When to use

Triggering scenarios:

- The user asks for an architecture diagram, flowchart, swimlane diagram, module relationship diagram, data flow diagram, or timing chain diagram
- The user asks to embed a diagram in a docx, or mentions "figure"/"illustration"/"Figure N-x"
- The user mentions Visio, vsdx, EMF, or vector graphics
- A batch of diagrams is needed for a technical proposal / overall design document / test outline / research report

Not applicable:

- Only a quick sketch is needed (inline SVG/widget is faster)
- Photorealistic bitmaps are needed (use an image generation tool)

---

## Prerequisites (COM route — do this before every job)

```bash
# 1) Is Visio installed?
ls "C:/Program Files/Microsoft Office/root/Office16/VISIO.EXE"
# 2) Environment self-check —— ★ check it all at once: Python version / pywin32 / whether Pillow can "really measure fonts"
#    (★ don't just do `import PIL`: it doesn't load the ImageFont extension, so it falsely passes under preview Python)
python -c "import sys; sys.path.insert(0,'scripts'); import visio_lib; ok,p=visio_lib.check_env(); sys.exit(0 if ok else 1)"
```

**★★ Python version requirement (learned the hard way)**

| Version | Usability |
|---|---|
| **3.9 – 3.12** | ✅ Recommended. Pillow 11.x is stable |
| 3.13 | ✅ Works (Pillow 11.x / 12.x both fine) |
| **3.15 alpha and other preview builds** | ❌ **Do not use**. Pillow 12.3 importing `ImageFont` crashes outright (`slot ID 85`), and only when a font is actually loaded |

★ **Why the prerequisite check must use `from PIL import ImageFont`**: the old version only did `import PIL`,
  but the `PIL` package's `__init__` does not load the `ImageFont` C extension — so under preview Python
  that step **passes**, and it only crashes later when `_font()` actually calls `ImageFont.truetype()`, which is expensive to debug.
  `visio_lib` now has the same self-check at the top level (see `check_env()`), which gives a clear version recommendation on failure.

If dependencies are missing, run `pip install pywin32 pillow`. To **avoid polluting the local environment**, use `uv` to start a temporary environment:

```bash
uv venv --python 3.12 .venv-vision && .venv-vision/Scripts/pip install pywin32 pillow
.venv-vision/Scripts/python scripts/selftest.py
```

**On an environment without Visio**: go back to the "XML backend quick start" at the top, or see `references/cross-platform.md`.

**PowerShell cannot instantiate COM** (blocked by security policy) — you must use Python. See `references/pitfalls.md` item 1.

---

## Workflow (COM route)

### Step 1: Fix the canvas size

First measure the body column width of the target document, and draw **at 1:1 scale** (don't draw at full A4 width, or the font size will be unreadable after scaling down).

```bash
python scripts/docx_embed.py 目标文档.docx      # prints e.g. "body column width 17.50 cm"
```

With no target document, design the canvas **150–175 mm** wide by default.

### Step 2: Write the drawing script

Create a `绘图/` directory in the project, one script per diagram. Copy the skeleton from `assets/example_flowchart.py`:

```python
import sys, os
sys.path.insert(0, "<skill>/scripts")
from visio_lib import VDraw

def main():
    # ★ exiting the with block (including on exception) auto-closes, so no VISIO.EXE is left behind; try/finally works too
    with VDraw(150, 108, theme="tech-cn") as d:   # canvas in mm
        d.lane(0, 0, 150, 74, "Processing chain")
        d.step_box(6, 12, 42, 26, "① Step 1", "description text")
        d.h_arrow(48, 25, 60)
        d.check()                                 # text-overflow self-check
        d.export_all("图表素材/图N1_某流程图")
```

Coordinate convention: **origin at the top-left corner of the page, Y axis pointing down, units in mm**.

Available primitives:

| Method | Purpose |
|---|---|
| `rect(x, y, w, h, **style)` | Rectangle |
| `oval(x, y, w, h, **style)` | Ellipse / circle (for interface number badges: a solid circle with `fill="white"` + a dashed outer ring with `fill=None, lpat=2`) |
| `text(x, y, w, h, s, size=, color=, align=)` | Text box |
| `line(x1, y1, x2, y2, arrow=4, lpat=1)` | Straight line (`arrow=4` solid arrowhead, `0` none) |
| `poly([(x,y),...], arrow=4, lpat=4)` | Polyline, arrowhead only on the last segment |
| `box(x, y, w, h, title, description)` | Title + description block |
| `step_box(x, y, w, h, title, description)` | Step box for flowcharts |
| `lane(x, y, w, h, label)` | Swimlane / zone (dashed frame + label on the left) |
| `h_arrow` / `v_arrow` | Horizontal / vertical arrow |
| `legend(x, y, w, [(text, lpat), ...])` | Legend row (entries are **evenly distributed** across the total width `w`; give `w` enough room — if it is too small or content runs past the canvas, export crops to the content and a blank area appears on the right out of nowhere, see §22) |
| `add_cp(shape, fx, fy)` / `side_cps(shape, "top bottom")` | Add connection points to a shape (**fy is measured from the top edge**); returns the row number |
| `connect(shapeA, cpA, shapeB, cpB, via=[...])` | **A single polyline connector**, both ends glued to connection points and routed through the via points; endpoints follow when a box moves |
| `connect(..., via=[...], smooth=10)` | **Curved connector**: via points act as spline control points, 10 samples per segment (8–16 is right). Arrow direction follows the tangent at the end; `connect(a, cp, (x, y))` lets you omit the end shape and write a free end |
| `rect/text(..., margin=0.4)` | Left/right inner padding inside a box (default 0.8 mm). When a box is only just wide enough for one line of text, tighten this instead of shrinking the font |
| `connect(..., label=, label_at=, label_gap=, label_side=)` | Connector **carries its own text**. ★ Labels always use **positioning relative to the line** (`label_at` defaults to the midpoint of the longest segment, `label_gap` is the perpendicular distance, `label_side` is which side) — **don't** hand-write absolute `label_pos` coordinates (they don't move with the line, so they drift when the layout changes, see §23) |

Style parameters use **theme field names** rather than hardcoded colors: `fill="muted_fill"`, `color="accent"`, `line="mgray"`.

Built-in themes (the `theme=` parameter):

| Theme | Use |
|---|---|
| `tech-cn` | Chinese engineering documents: black/gray lines + dark-blue titles + gray background for external objects (**default**) |
| `mono` | Strict black-and-white printing |

### Step 3: Self-check → fix → export

**Visio does not report an error when text overflows its box** — it just lets the text spill outside the box. You must self-check:

```python
d.check()
# [Text overflow] 2 places
#   模板管理  8.5pt box width 34.0 box height 26.0 → 4 lines need 31.2mm
# [Connector check] exit direction / crossing solid boxes: none
```

`d.check()` also checks **connectors** (since 2026-10-17):

- `[Connector direction] … glued to the left side of shape #N, but the first segment heads the opposite way` — the connector exits the box from the **wrong side**,
  so it runs straight across the whole box (hidden underneath when the box has a fill, nearly invisible to the eye). `connect()` prints this line as it writes.
- `[Connector crossing] line #N enters shape #M … interior` — a polyline crosses a **filled** solid box (containers such as swimlanes/border frames and fill-less full-width text boxes don't count). See `references/pitfalls.md` §24.

Fixes: **first make the box taller, then cut the text**, and only consider shrinking the font last (a shrunken font is unreadable once scaled down on embedding).
Rule-of-thumb sizes: box titles 8.5–9 pt, in-box descriptions 7–7.5 pt, left/right inner padding 0.8 mm.

After exporting, also check whether **arrows have drifted** (an arrow is attached to a box but lands off the midpoint of the box edge, so it looks crooked):

```bash
python scripts/audit_align.py 图表素材/图N1.vsdx
```

It reports endpoints that are "attached to one edge alone, yet more than 2 mm off that edge's midpoint". This is a **heuristic screen**:
if a box fans several arrows out of the same side (each reaching a box at a different height), the offset is normal and needs a visual check against the rendered image.

Endpoints **confirmed as "deliberate offsets"** can be registered in the whitelist (`scripts/audit_whitelist.json`) so batch regression no longer counts them as failures:

```bash
python scripts/audit_align.py --add 图表素材/图N8….vsdx   # register the offsets reported this time in the whitelist
python scripts/audit_align.py 图表素材/*.vsdx            # afterwards, confirmed diagrams → whitelist-exempt, exit 0
```

The match key is the diagram name (without extension) + endpoint coordinates (±3 mm tolerance) + edge name; a hit prints "→ whitelist (the note recorded at registration)";
if an entry no longer matches anything in the diagram (layout changed / already fixed) → a closing hint says "expired, consider cleaning it up".
★ Registering = a promise to skip the check, so **confirm by hand first, then --add** — don't use it to silence real defects.

> To see the deviation directly, you can overlay the vsdx's exact geometry as red lines on the exported PNG for comparison —
> red and black lines coinciding means the geometry is right; a mismatch means the coordinates were written wrong (see `references/pitfalls.md` §11).

Iterate until `[Text overflow] none` and the arrow-alignment audit passes, then finalize. See `references/pitfalls.md` items 3, 4, and 11.

### Step 3.5: Merging text into its host (automatic by default)

"Rectangle + standalone text box" is a common pattern, but **a rectangle can hold its own text**, and a connector carries its own `Text` too.
`VDraw` automatically calls `merge_texts()` before `save()` / `export_all()`, merging
**text boxes that fall inside a rectangle** into the rectangle, and **small labels sitting on a connector** into the connector, with the **layout pixel-identical**
(12-diagram regression: per-tile median displacement 0.00 px, p95 1 px). Shape count drops by ~47%.

```python
d.merge_texts()                  # usually no manual tuning needed; save()/export runs it automatically
d.merge_texts(verbose=True)      # prints: [Text merged] hosts 9 / merged text boxes 15 ｜ connector labels 3 …
```

Switches and debugging:

| Method | Effect |
|---|---|
| `VISIO_MERGE=0` environment variable | Disables auto-merge → produces an "unmerged" comparison diagram for pixel-by-pixel regression |
| `VDraw.default_auto_merge = False` | Same, in code (subclasses can override) |
| `VISIO_MERGE_DEBUG=1` environment variable | Prints where each text box went and the hit distance of connector labels |

Rules: a text box **fully inside** some rectangle → merged (if nested, take the smallest one); skipped if the rectangle already has text;
a connector label must be **≤44 mm wide and ≤12 mm tall** (so zone titles/legends/description paragraphs are not swallowed by mistake),
and its center must be within 6 mm of the polyline. **After merging, `Text` becomes rich text** (per-run font size/color/space-before restored),
so to change wording, edit the source script and rerun — don't hand-edit in Visio.

> How it works and the pitfalls (rich text `CellsSRC(3,r,7)`, true block height `_BH_EXTRA`, PIL font-size rounding undercounting line wraps,
> connector `TxtLocPinX` defaulting to 0, local origin ≠ Begin point) are in `references/pitfalls.md` §18.
>
> ★ **If you want a connector to carry its own label, don't rely on the heuristic above to guess — pass `label=` explicitly**:
>
> ```python
> d.connect(a, cp["a"]["right"], b, cp["b"]["left"],
>           label="Dual 10GE", label_color="gray",
>           label_at=0.5, label_gap=2.2, label_side="up")   # positioned relative to the line, moves with it
> ```
>
> Always position with `label_at` / `label_gap` / `label_side` (positioning relative to the line);
> **don't** hand-write absolute page coordinates in `label_pos` — they don't move with the line, so a layout change leaves the label adrift.
>
> A 1-D connector's own text **doesn't draw a big rectangle when selected** and moves with the line; loosening the heuristic instead merges several text boxes onto the same line, where they overlap and lose text. See `references/pitfalls.md` §20.

### Step 4: Embed in Word

**Recommended route A** — leave a placeholder in the document and batch-replace after drawing (Word natively supports EMF, so vectors stay crisp):

```markdown
{{FIG:N2}}
```

```bash
python scripts/word_insert.py 方案.docx --placeholder "{{FIG:N2}}" --image 图N2.emf
```

Default width = body column width − 5 mm; height is computed automatically from the EMF aspect ratio.

**Route D** — the document **already contains** a Visio OLE diagram to be replaced in place (rework / editing a client-supplied document):

```bash
python scripts/replace_ole.py 方案.docx 图表素材/图N2.emf 图表素材/图N2.vsdx
```

This replaces `word/embeddings/*.vsdx` and `word/media/*.emf` together (Word normally draws the cached EMF,
so replacing only one side produces "the displayed diagram ≠ the diagram you get on double-click"), and rewrites the
display frame from the EMF's `rclFrame` aspect ratio (default width = body column width − 0.5 pt). See `pitfalls.md` §31 for details.

**Route B** — when building a docx from scratch, hand-assemble the OOXML; see `references/docx-embed.md` (including three must-hit pitfalls).

**Route C** — when the client requires "double-click to edit in Visio", swap the already-inserted diagram **in place for an OLE object**:

```python
rng, w, h = sh.Range, sh.Width, sh.Height   # first take the Range and the "width + height"
sh.Delete()
new = rng.InlineShapes.AddOLEObject(FileName=abs_vsdx)
new.LockAspectRatio = 0                     # must turn this off first, or a later write to one dimension drags the other
new.Width = w                               # restore the width by hand
new.Height = h                              # ★ the height must also be restored by hand
```

> ★ **Writing only `Width` destroys the frame**: `AddOLEObject` sizes by the **OLE object (Visio)'s own intrinsic aspect ratio**,
> and Word then recalculates the height itself, which doesn't match the EMF frame — measured diagrams were squashed by **up to +20%** (boxes visibly flattened).
> The height must be taken from the `sh.Height` of the "diagram being replaced" (the only trustworthy ratio, since it was computed from the EMF frame) and written back explicitly.
> See `references/pitfalls.md` §13.

Key points: `InlineShapes` correspond one-to-one with the diagram list in document order; what gets embedded is the `.vsdx` (so keep the `.vsdx`);
after replacing, run `TablesOfContents.Update()` to reflow page numbers; finally remove `updateFields` from `settings.xml`.
For the complete code and checks, see `references/docx-embed.md`, route C.

### Step 5: Pre-delivery check

```bash
python scripts/docx_audit.py 方案.docx
```

Static checks (no Word needed): XML well-formedness, Content_Types declarations, `mc:Ignorable` prefix consistency, `tblW`/`tcW` attribute names, leftover bold, body column width.

Live Word checks: page count, **widest table vs column width**, each diagram's size, bold counts inside and outside tables.

Pass criteria: widest table ≤ body column width; bold inside tables = 0; bold outside tables = number of heading paragraphs (heading styles are bold by themselves, which is normal).

---

## Key constraints (must follow)

1. **Draw at 1:1** — canvas width ≈ body column width, or fonts will be unreadable after embedding.
2. **Always self-check** — `d.check()` is a fixed step of the export flow; `TextBounds` doesn't exist, so there is no other way to catch overflow.
   It also outputs a **connector self-check** (exit direction / crossing solid boxes); both must read "none" to pass.
3. **Always close (`d.close()`)** — otherwise VISIO.EXE lingers as a background process, and accumulated instances make COM calls time out.
   Prefer `with VDraw(...) as d:` (exits auto-close, including on exceptions), or close in `try/finally`;
   `export_all()` now **raises on failure** (no longer silently swallows it), so don't forget the finally fallback.
4. **Colors via theme field names** — don't hardcode `"RGB(31,78,121)"`, so the theme can be swapped as a whole.
5. **Include the gutter when computing column width** — ignoring `w:gutter` makes tables/diagrams overflow the page.
6. **Connectors: use `connect()` as a single polyline + glued connection points** — don't piece together multiple `poly()` segments (they don't follow when a box moves and can't be edited as a whole in Visio); and don't use `Page.Layout()` to rearrange (it moves boxes you already placed).
7. **Run the self-test after changing the library** — `python scripts/selftest.py` (22 regression assertions, including offline EMF header checks); for documentation changes run `python scripts/doc_consistency.py` (guards against conclusions drifting). **Give the legend's total width `w` enough room**, or the whole row runs past the canvas (§22).

---

## Known limitations (know before you use it)

1. **Free-end connectors (one end unglued) still show a bounding box when selected** — legend lines / tick marks / arrows / bus stubs:
   if such a line is given `ObjType=2`, Visio clamps it to the 1-D minimum size after `SaveAs` and **breaks the geometry** (tick marks get shortened,
   stubs gain a slanted segment). This library takes "**correct geometry first**": such lines get **no `ObjType`**, at the cost of still drawing a box when selected.
   Connectors glued at both ends are unaffected (normally no box). See `pitfalls.md` §21.
2. **Text merging (`merge_texts`) is heuristic** — it judges by "is the text box inside a rectangle / how far is the label from the line",
   so edge cases may swallow or miss text. **For precise control use an explicit `label=`** (written as the connector's own text); don't rely on the heuristic.
3. **Depends on local Visio + pywin32; not purely offline-portable** — PowerShell cannot instantiate COM (sandbox policy),
   so you must use Python; a new machine needs Visio plus `pywin32` / `pillow`, and CJK fonts (default "Microsoft YaHei").
4. **Arrowheads only come in "medium"** — shapes produced by `DrawLine`/`DrawPolyline` have no `ArrowSize`
   cell (`pitfalls.md` §30), so arrowheads don't scale with line weight. In tight spots (interface badges within 3 mm of each other),
   solve it with "enough length + near-tangent direction + visual check at 96 dpi" — don't expect `LineWeight` to shrink the arrowhead.
5. **Curves are dense-polyline approximations** — `connect(smooth=)` samples Catmull-Rom (`pitfalls.md` §30),
   so there are no draggable control points in Visio; `Page.DrawSpline` simply doesn't work under COM, so don't try it.
6. **COM is a single instance; the process must be closed by you** — forgetting `d.close()` leaves a `VISIO.EXE` behind, accumulated instances make COM calls time out,
   and a leftover process makes Word's OLE embedding **reuse the old document** (`pitfalls.md` §14).

---

## Resources

### scripts/

| File | Purpose |
|---|---|
| `visio_lib.py` | Drawing engine: the `VDraw` class (supports `with`, auto-close on exit) + `text_extent_mm()` font metrics + `check_env()` environment prerequisite check |
| `xml_backend/` | **Cross-platform XML backend**: `vlib.py` drawing DSL → editable vsdx + PNG; `build_vsdx.py` / `build_docx.py` / `extract_vsdx.py` / `inspect_visio_docx.py` / `visio_finish.py` (Windows verification/finish). See `references/cross-platform.md` |
| `font_support.py` | Cross-platform default fonts (Windows=Arial+Microsoft YaHei / macOS=Helvetica Neue+PingFang / Linux=fontconfig), shared by both routes; `VISIO_FONT_*` environment variables override |
| `theme.py` | Three color themes + font file mapping |
| `docx_embed.py` | `emf_size_mm()` reads the EMF frame · `content_column_mm()` computes the body column width (**defaults to the last section, supports multi-column conversion**) · `ensure_content_type()` adds missing declarations |
| `word_insert.py` | Route A: Word COM inserts diagrams by placeholder (with command line) |
| `replace_ole.py` | Route D: replaces a Visio OLE diagram already in the document in place (vsdx + cached EMF + display frame together) |
| `audit_align.py` | Arrow-alignment audit: finds "arrow attached to a box but landing off the edge midpoint" drift; **confirmed deliberate offsets can be registered with `--add` in the whitelist** (`audit_whitelist.json`, ±3 mm coordinate matching, expiry hints) |
| `docx_audit.py` | Pre-delivery self-check: static OOXML checks + live Word layout measurement |
| `selftest.py` | **Library self-test**: draws a test diagram covering 15 regression categories A–O → asserts ObjType / geometry point counts / connector labels / text merging / legend not overflowing / EMF header rewrite (**run it after changing the library**, `python scripts/selftest.py`) |
| `doc_consistency.py` | **Doc–code consistency**: deprecated-phrasing blacklist + key conclusions must be present + paired constraints, guarding against "code changed, docs still state the old conclusion" drift |

### references/

| File | Contents |
|---|---|
| `cells-reference.md` | Visio CellsU property quick reference, coordinate conversion formulas, numbering values |
| `pitfalls.md` | 29-item pitfall checklist (COM sandbox, TextBounds doesn't exist, overflow, 1:1, gutter, EMF dimensions, COM busy errors, arrow drift, bend points pressing on neighboring boxes, **OLE frame squashed**, **leftover VISIO.EXE makes OLE embed old content**, **connection-point Y pointing up flips fy**, **polyline geometry must be translated into the bounding box**, **single-polyline connector recipe**, **merging text into its host**, **endpoint side choice for vertically adjacent boxes**, **connector-carried text label=**, **rectangular outline when a polyline is selected (root cause: `Misc.ObjType` not set to 2)**, **`legend()` ignoring w leaves a big blank area on the right**, **connector labels far from the line — absolute coordinates / merging not snapping to the line**, **diagram squashed / wrong aspect ratio — first do the three-way comparison (display frame · bounds · vsdx union)**, **Word save normalizing 72pt to 1in**, **static EMF diagram text clipped, only a full redraw fixes it**, **the centering / table-of-contents page-number measurements themselves are the pitfall**…) |
| `docx-embed.md` | Both embedding routes + complete fix code for the three OOXML pitfalls |
| `cross-platform.md` | Full XML backend description: division of labor with COM, usage, fonts, known limitations |
| `visio-style-rules.md` | Diagram style rules and review checklist (layout / connectors / font sizes / colors) |

### assets/

| File | Contents |
|---|---|
| `example_flowchart.py` | Minimal runnable example (~60 lines drawing a flowchart with a loop and a legend); copy it and change the content |
| `example_xml_flowchart.py` | Minimal XML backend example (cross-platform, no Visio needed); copy it and write FIGS |

---

## Quick troubleshooting

| Symptom | Cause | Where |
|---|---|---|
| On Linux/macOS, `import visio_lib` fails with missing pywin32 | You took the COM route; switch to the XML backend (see the quick start at the top) | `cross-platform.md` |
| XML backend can't find CJK fonts | Install CJK fonts for the system, or set `VISIO_FONT_CJK_FILE` | `cross-platform.md` §Fonts |
| COM instantiation refused | PowerShell policy blocks it | `pitfalls.md` §1 |
| `TextBounds` AttributeError | That API doesn't exist | `pitfalls.md` §2 |
| Text runs outside its box after export | Visio doesn't report it; self-check required | `pitfalls.md` §3 |
| Arrow looks crooked / nearly touching a box corner | Landing point is off the edge midpoint; check with `audit_align.py` | `pitfalls.md` §11 |
| Text too small after embedding | Canvas was drawn at full A4 width | `pitfalls.md` §4 |
| Table overflows the right side of the page | Gutter not counted | `pitfalls.md` §5 |
| Word reports "file is corrupted" | Content_Types or mc:Ignorable | `docx-embed.md` pitfalls 1 / 2 |
| Page count explodes, one character per column | `tblW`/`tcW` written as `w:val` | `docx-embed.md` pitfall 3 |
| Wrong size after embedding | Trusted the EMF's absolute dimensions (only the aspect ratio should be used) | `pitfalls.md` §9 |
| **Diagram squashed after OLE replacement** | **Only `Width` was written, so Word recalculated the height from the OLE's own ratio** | **`pitfalls.md` §13** |
| `Call was rejected by callee` | Word COM is busy; retry + make cleanup non-throwing | `pitfalls.md` §10 |
| VISIO.EXE process left behind | Forgot `app.Quit()` | `pitfalls.md` §7 |
| **One diagram still shows old content after OLE replacement** | **Leftover VISIO.EXE; the OLE server reused the old document** | **`pitfalls.md` §14** |
| Vertical line runs through two boxes / connection point lands on the wrong side of a box | `Connections.Y` is a local coordinate with Y pointing up, so `fy` must be measured from the top edge | `pitfalls.md` §15 |
| Blank rectangle when a polyline is selected / selection box not on the line | Geometry rows weren't translated into `[0,Width]×[0,Height]`; pair with `LocPin` to pull the pin back | `pitfalls.md` §16 |
| **Rectangular outline along the bounding box when a polyline is selected** | **`Misc.ObjType` wasn't set to `2` (routable 1-D shape). Must do `ConFixedCode=3` → `ObjType=2` → **read back the local origin** and rewrite the geometry (order-sensitive); short shapes / free ends must fall back** | **`pitfalls.md` §21** |
| **Whole diagram squashed / wrong aspect ratio** | **Display frame was computed from the EMF `frame` (whose X/Y DPI differ, about 1.914×). First do the three-way comparison: display-frame ratio / bounds ratio / vsdx union ratio, then decide what to fix** | **`pitfalls.md` §25** |
| Audit reports "display frame 340x0" / "aspect ratio −100%" | Word save normalizes `72.0pt` to `1in`; style parsing must support pt/in/mm | `pitfalls.md` §26 |
| **A box title is missing a character / text presses against the frame (static EMF, no vsdx source)** | **Text in EMF is vector outlines, so PDF extraction can't reach it; metric-based thresholds are unreliable. Only a full redraw fixes it** | **`pitfalls.md` §27** |
| Diagram not centered / table-of-contents page numbers don't match | The measurement methods themselves have pitfalls (`Range.Information(5)` is invalid; footer blocks can't be taken in row order) | `pitfalls.md` §28 |
| Vertical line runs through the box above/below | For vertically adjacent boxes, the connector must go from the lower box's `top` to the upper box's `bottom` | `pitfalls.md` §19 |
| **Big rectangle when a connector label is selected** | The label is a 2-D text box; it should be the connector's own text (`connect(..., label=)`) | **`pitfalls.md` §20.0** |
| Connector label jumps / rotates with the line | `Control.TextPosition` was written (Visio interprets it in page coordinates) | `pitfalls.md` §20.2 |
| **Two-character label squeezed into "one character per line"** | The connector text block has left/right page margins and `TxtWidth` is too small → zero the margins and leave headroom | **`pitfalls.md` §20.3** |
| **A connector label's whole text disappears** | Loosening the `_merge_line_labels` heuristic merged several text boxes onto one line, where they cover each other | **`pitfalls.md` §20.4** |
| **Connector label far from its line** (especially after a layout change) | An absolute `label_pos` was written (doesn't move with the line), or the label is a standalone text box whose offset was merged into the connector verbatim. Use `label_at/label_gap/label_side` positioning relative to the line; the merge path now snaps to the line automatically | **`pitfalls.md` §23** |
| **Big blank area on the right of the diagram / legend entries unevenly spaced** | `legend()`'s `w` was too small (or a legend/long text ran past the canvas), and export cropped to the content including the overflow | **`pitfalls.md` §22** |
| **Page still shows the old diagram after replacement / only the double-click view is new** | Only the vsdx was replaced, not the `word/media/*.emf` cached preview | `replace_ole.py` / `pitfalls.md` §31 |
| **Diagram squashed after replacement** | The `v:shape` display frame wasn't rewritten to the new EMF's `rclFrame` aspect ratio | `pitfalls.md` §31.2 |
| `DrawSpline` reports `Invalid number of parameters` | That API doesn't work under COM | `pitfalls.md` §30.1 (use `smooth=`) |
| Arrowheads too big / crowded together | The drawn shape has no `ArrowSize` cell, so the size is fixed at medium | `pitfalls.md` §30.2 |
| A `/skill-name` command loads the **old** content / base directory points at `*.bak_*` | A **backup copy** of the skill got mixed into the skills directory (the scanner treats it as a candidate). Move backups out of skills/ (e.g. `~/.workbuddy-ai/skill-backups/`) and keep only the original | This item |
