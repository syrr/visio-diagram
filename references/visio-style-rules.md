[中文文档 (Chinese)](visio-style-rules.zh-CN.md)

# General Guide to Modifying Embedded Visio Objects (for LLMs / Agents)

Version: v2.3 (2026-11-19)
Scope: any Visio object (.vsdx OLE object) embedded in a Word (.docx) document, and standalone .vsdx artwork. Not tied to any specific project.
Audience: LLMs / agents responsible for inspecting, unifying the style of, redrawing, or modifying such figures, as well as human editors.
How to use: read §1–3 first to settle the workflow; follow the visual rules in §4–9 when modifying figures, the structural rules in §10 when touching files; tools are in §11–12, and the pre-delivery checklist is §13. The user's explicit requirements take priority over this guide; the guide's defaults (page width, palette, fonts) may be replaced as a whole per document or per user request, but only one set may be used within a single document.

> Project-specific information (figure-number-to-embedding-file mapping, the layout choice for each figure, modification records) does not go in this file; it goes in the project's own ledger file (see §14 for the recommended approach).

---

## 1. General Principles

1. **Content unchanged, form unified.** Keep all entities, relationships, and text of the original figure; change content only to fix clear errors (text/figure mismatch, dangling arrows, duplicate numbering), and record each change.
2. **One figure, one idea.** Each figure answers a single question (composition, flow, timing, correspondence); don't mix drawing styles.
3. **Consistent with the body text.** Names, numbers, and flow order in the figure match the body text and the figure caption; when they disagree, change the figure to align with the text first, and record it.
4. **Stay editable.** The figure must remain an embedded object that can be edited in Visio by double-clicking; never replace it with a plain picture.
5. **Preview in sync with source.** What Word shows is the object's preview image, not the vsdx itself; after changing the vsdx you must also update the preview, otherwise the user still sees the old figure.
6. **Reproducible.** Modifications are generated from the source manuscript by script; the source is read-only; anyone re-running the script gets the same result.

## 2. End-to-End Workflow

1. **Back up and unpack**: save the source manuscript as a read-only copy; a docx is a zip, so unpack it into a temp directory.
2. **Inventory**: run the diagnostic script (§11.3) to list each object's embedding file, preview file, page size, Word display size, font sizes, and connector problems; build a "figure number → embedding file" cross-reference in document order and write it into the project ledger.
3. **Visual inspection**: render each figure's preview as a PNG and check every one with an image viewer for overlaps, overflow, dangling ends, and text pressed against lines.
4. **Issue a diagnostic report**: first give the user "global issues + per-figure issues + recommendations", and only modify after getting agreement.
5. **Set the spec**: settle page width, fonts, and palette (defaults in §4–8 of this guide).
6. **Redraw figure by figure**: describe each figure with the layout DSL (§12), choosing one layout template per figure (§9).
7. **Write back**: replace the vsdx page XML, page size, preview image, and the object size in Word (§10).
8. **Self-check**: XML parses, the diagnostic script reports no issues, every preview passes visual inspection (§13).
9. **Deliver**: send the revised manuscript to the user, explaining what changed and what limitations remain (e.g. no drag test done in real Visio, which connectors are unglued by design).
10. **Record**: append modification records (date, version, per-figure changes) to the project ledger.

## 3. Diagnostic Points (Common Problems and How to Judge Them)

| Problem | How to spot it | Handling |
|---|---|---|
| Font size too small and inconsistent across figures | In-page font size × Word scale = effective font size; < 7.5 pt or large differences across figures | Unify page width = text-block width, insert at 100%, unify font sizes |
| Non-uniform stretching | Word display width / page width differs from display height / page height by > 2% | Reset to 100% and lock the aspect ratio |
| Scaled a second time | Word display width ≠ page width | Change page width to the text-block width, then insert at 100% |
| Diagonal segments | Adjacent geometry points of a connector change in both x and y | Fix to orthogonal routing |
| Dangling arrows / endpoints not following on drag | Connector has no `Connect` record; or has a record but BeginX/EndX have no glue formula; or the endpoint is not on a shape edge | Drop endpoints onto shape edges; write connection points, glue formulas, and Connect per 7.2 |
| Connector text as a standalone text box | A small borderless text shape near a connector, containing "Pass/Send" etc. | Delete the text box; write the text into the connector's own `<Text>` (7.3) |
| Branch line starting from the middle of a line | Branch start lands on another line, not on any shape edge | Change it to start from the source shape's edge, overlapping the main line along the same segment (7.2) |
| Rounded-corner abuse | `Rounding` ≠ 0 on anything that is not a flow terminator | Change to 0 |
| Legend / explanatory text on the figure | "Note:" or legend dashes inside the figure | Delete; put explanations in the body text; express relationships with nodes |
| Font formula errors | `#VALUE!` / `#REF!` / `#NAME?` appearing in page XML | Write Font and AsianFont on every Character row |
| Connector text pressed against a line or a border | Visual inspection | Put the text at the midpoint of the longest segment, on a white background (7.3) |
| Text overflowing / squeezed | Visual inspection; estimated text width > box width | Widen the box or break lines by meaning |
| Isolated shape | No connectors at all, and not a container/explanation | Change to an auxiliary style or add a relationship |
| Missing-glyph symbols | ▾, ☐ etc. rendering as boxes | Replace with text or indentation |

Note: lines whose endpoints by nature have no shape there — timelines, tick marks, separators — may be ignored in the "unglued connector" diagnosis; all other unglued connectors must be handled.

## 4. Layout and Dimensions

| Item | Rule |
|---|---|
| Page width | Equal to the Word text-block width (commonly 150–170 mm for A4); the same for all figures in a document |
| Page height | Set by content, in multiples of 5 mm; commonly 80–150 mm |
| Insert into Word | **At 100% scale**, width = page width (mm)/25.4×72 pt, height = page height (mm)/25.4×72 pt, aspect ratio locked |
| Outer margins | Content ≥ 4 mm from page edges |
| Grid | Coordinates and sizes aligned to 0.5 mm; same-type elements aligned to the same baseline |
| Element spacing | 2.5–8 mm between boxes in the same row; 8–10 mm between layers (containers) |
| Forbidden | Non-uniform stretching; second scaling in Word causing font-size drift |

## 5. Fonts and Font Sizes

| Use | Font | Size | Weight | Color |
|---|---|---|---|---|
| Ordinary box text | Chinese documents: Microsoft YaHei (Latin name Microsoft YaHei); English documents may use Segoe UI / Arial | 9 pt | Regular | #1A1A1A |
| Key box text | Same as above | 9 pt | Bold | #1F3864 |
| Container title (title band) | Same as above | 9 pt | Bold | #FFFFFF |
| In-figure title | Same as above | 10 pt | Bold | #1F3864 |
| Connector text (decision branches only) | Same as above | 8 pt | Regular | #404040 |
| Floor in crowded spots | — | No less than 7.5 pt (counted as the effective size in Word) | — | — |

- Write `Font` and `AsianFont` in every shape's Character section; leave no formula errors.
- Break multi-line text by meaning, each line complete, not overflowing the border.
- Don't use symbols missing from the target font; show tree hierarchy with indentation.
- When rendering previews without Microsoft YaHei, substitute metrically close Noto Sans SC (see 11.2).

## 6. Colors and Shape Styles (Default Palette)

Use only one palette per document; don't introduce colors outside the palette. Default engineering-style palette (table below), plus 1 optional theme (mono), defined in `XML_PALETTES` in `scripts/theme.py` (sharing a registry with the COM themes, same name means the same thing on both routes):

| Theme | Suited for | Primary color |
|---|---|---|
| `tech-cn` (default) | Chinese engineering documents | Dark blue #1F3864 |
| `mono` | Strict black-and-white printing | All grayscale; ok/warn distinguished only by fill darkness |

Usage: single figure `Fig(W, H, theme="mono")`; batch `build_vsdx.py / build_docx.py --theme mono` (or the `VISIO_THEME` environment variable). Changing themes changes only colors: the semantic roles, layout, font sizes, and connector rules of `node/key/dark/ext/ok/warn/soft` stay the same; ok/warn are identical across themes except mono; the no-shadow, no-gradient, no-3D rules apply to all themes. The default remains tech-cn; omitting theme leaves the appearance unchanged.

| Style name | Fill | Border | Line type | Text | Use |
|---|---|---|---|---|---|
| node ordinary box | #FFFFFF | #2E5C8A 0.75 pt | Solid | #1A1A1A | General modules, steps |
| key key box | #DCE8F5 | #2E5C8A 0.75 pt | Solid | #1F3864 bold | Core modules, main-chain nodes |
| dark critical box | #1F3864 | #1F3864 | Solid | #FFFFFF bold | The 1–2 most critical nodes per figure |
| ext external box | #F2F2F2 | #7F7F7F 0.75 pt | Dashed | #404040 | External systems, objects under test, external inputs |
| soft auxiliary box | #F3F6FA | #A6B9D0 0.5 pt | Solid | #1F3864 | Explanation bars, environment, secondary information |
| ok positive result | #E2F0D9 | #548235 | Solid | #385723 | Match, pass |
| warn exception result | #FBE5D6 | #C55A11 | Solid | #843C0C | Deviation, missing, rejected |
| container container | #F7F9FC | #8EA9C1 0.75 pt | Solid | — | Layers, groups |
| container-ext external container | None | #8EA9C1 | Dashed | — | External area |
| band title band | #2F5597 | #2F5597 | — | #FFFFFF bold | Container's top-left title, 6.5 mm high |
| lane header swimlane header | #1F3864 | — | — | #FFFFFF bold | Left-side swimlane title, about 22 mm wide |

- Rounded corners: rectangles are always square (`Rounding=0`); only the start/end terminators of flowcharts use rounded corners (capsule shape).
- Keep it simple: no legends, annotations, or explanatory text on figures; connectors carry a condition label only on decision branches (e.g. "Pass/Fail"). Meanings previously carried by connector labels (path names, data content) become nodes instead; explanatory text goes in the body. Figure numbers and names appear only in the body-text caption.
- No shadows, no gradients, no 3D.
- dark: at most 2 per figure; ok / warn only for classifying results, not decoration.

## 7. Connectors

### 7.1 Line Types and Routing
| Type | Color | Line type | Line width | Use |
|---|---|---|---|---|
| Main data flow / control flow | #2E5C8A | Solid | 1 pt | Forward flow |
| Return, archiving, reporting, offline | #7F7F7F | Dashed | 1 pt | Reverse or non-real-time data |
| Closed-loop iteration | #2F5597 | Dashed | 1 pt | Iteration back to the start |
| Exception branch | #C55A11 | Dashed | 1 pt | Rejected, failed branch |

- Always **orthogonal (right-angle) routing**, no diagonal segments; at most 3 bend points.
- Arrows: solid triangle (Visio `EndArrow=13`, size 1).
- Don't cross unrelated shapes, title bands, or text; route around through gaps between layers.
- For multiple lines merging/branching, use a shared horizontal "bus" segment, with branch points in the gaps between layers.
- When a one-to-many support relationship is hard to draw, use a single "group-to-group" relationship (e.g. support-module-group container → processing-chain container) instead of multiple dangling arrows.

### 7.2 Endpoint Gluing (Endpoints Follow When a Shape Is Dragged)
Goal: when the user drags any shape in Visio, the endpoints of connectors attached to it move along. Writing `<Connects>` alone is not enough — Visio implements the follow behavior with formulas on the endpoint cells.

1. **Every bordered shape (rectangle, container) carries connection points**: `<Section N="Connection">`, by default the four edge midpoints, plus extra connection points wherever a connector actually lands. Write a formula on each row so the connection points follow when the shape is resized:
   ```xml
   <Section N="Connection">
     <Row IX="0"><Cell N="X" V="0.59" F="Width*0.5"/><Cell N="Y" V="0.47" F="Height*1"/>
       <Cell N="DirX" V="0"/><Cell N="DirY" V="0"/><Cell N="Type" V="0"/><Cell N="AutoGen" V="0"/></Row>
     …
   </Section>
   ```
   Row IX is called `Connections.X{IX+1}` / `Connections.Y{IX+1}` in formulas. The Connection section goes after Character and Paragraph, before Geometry.
2. **Write glue formulas on connector endpoints** (point-to-point gluing, keeping the original exit position):
   ```xml
   <Cell N="BeginX" V="1.23" F="PAR(PNT(Sheet.3!Connections.X2,Sheet.3!Connections.Y2))"/>
   <Cell N="BeginY" V="2.10" F="PAR(PNT(Sheet.3!Connections.X2,Sheet.3!Connections.Y2))"/>
   ```
   Same for EndX/EndY. `Sheet.n` is the target shape's Shape ID, so output all 2-D shapes first, then the connectors.
3. **Also write `<Connects>`**, one-to-one with the formulas:
   ```xml
   <Connect FromSheet="20" FromCell="BeginX" FromPart="9" ToSheet="3" ToCell="Connections.X2" ToPart="101"/>
   ```
   `FromPart`: 9 for the start point, 12 for the end point; `ToPart` = 100 + connection-point row IX.
4. **The connector's own size follows its endpoints by formula** (no rotation, Angle=0): `PinX=(BeginX+EndX)/2`, `PinY=(BeginY+EndY)/2`, `Width=EndX-BeginX`, `Height=EndY-BeginY`, `LocPinX=Width*0.5`, `LocPinY=Height*0.5`. The local coordinate origin is the start point, the end point is at (Width, Height), and width/height may be negative or 0.
5. **Write a formula for every point of a polyline's geometry**: X as `Width*(x/W)`, Y as `Height*(y/H)`. In a direction where W or H is 0 (X of a purely vertical line, Y of a purely horizontal line), a bare number is not enough — otherwise, when the target is dragged sideways, the endpoint follows but the line body stays put: write `Width*0` for the first point and `Width*1` for the last point, and have middle points follow the nearer end by index while keeping their original offset (e.g. `Width*0+5 mm`); `TxtPinX/Y` likewise, following the midpoint when the span is zero. This way L- and Z-shaped polylines stay right-angled after endpoints move, and a line turning diagonal doesn't detach from its endpoints.
6. **Make it a routable connector**: `ObjType=2`, `ShapeRouteStyle=1` (right-angle), `ConLineRouteExt=1`, `ConFixedCode=0`. Visio may adjust bend-point positions when re-routing; that's normal.
7. **Endpoints must land exactly on shape borders** (tolerance ≤ 0.3 mm), otherwise no glue target is found. When auto-detecting, take the smallest shape whose border contains the point; don't glue to a pure text box.
8. **Don't branch off the middle of a line**: a line's midpoint is not a shape and cannot be glued to. Change the branch to start from the source shape's edge, with its first segment overlapping the main line along the same path; the appearance is unchanged and both ends can be glued.
9. Leave unglued only when there really is no shape at the endpoint: timelines, tick marks, separators, and the end of an arrow pointing at a separator.

### 7.3 Connector Text
- Connector text is written in the **connector's own text**, not a separate text box; that way the text moves with the line when the connector or a shape is dragged, and is deleted together with the connector.
- How to write it: in the connector Shape add `TxtPinX/TxtPinY` (formulas proportional to Width/Height; for the zero-span direction see the previous item; place at the midpoint of the longest segment), `TxtWidth/TxtHeight`, `TxtLocPinX=TxtWidth*0.5`, `TxtLocPinY=TxtHeight*0.5`, `TxtAngle=0`, a white text background `TextBkgnd` (`F="RGB(255,255,255)"`), `TextBkgndTrans=0`, and zero text margins; then add Character (8 pt) and Paragraph (centered) sections, and finally `<Text>`.
- Use it only on decision branches (meanings beyond 7.1 become nodes).

## 8. Titles, Numbering, and Explanations

- Figure numbers and names come from the Word caption; don't repeat the figure number inside the figure.
- No legends or on-figure annotations; explain what line types mean in the body text, or express them with nodes instead.
- Number steps with ①②③…; only one numbering set per figure.
- Figure terminology matches the body text: when the body text's wording changes (e.g. some function changes from "store and look up" to "infer by model"), the node names in the affected figures change in step, and are recorded.

## 9. Layout Templates (by Figure Type)

| Figure type | Layout |
|---|---|
| Layered architecture diagram | Top-down container layers (e.g. Client → Services → Runtime environment), equal-width modules in a horizontal row within each layer, no labels on inter-layer arrows |
| System composition / network diagram | This system in a left/top container, external objects as ext on the right/bottom; every terminal device must have a connector joining the network, no isolated devices left |
| Linear flowchart | A single column or row of equal-width step boxes, main line centered; loops run along the outside; only start/end use rounded corners |
| Decision / branching flow | Input → decision layer → result layer (ok/warn) → summary → output, branches on a bus; branch conditions written in connector text |
| Two-party interaction diagram | Two equal-width containers left and right, boxes in the same row corresponding one-to-one, horizontal arrows crossing the middle seam |
| Swimlane / correspondence diagram | Horizontal swimlanes + dark swimlane header on the left; or multi-column headers + row alignment |
| Timing diagram | Segmented containers; the timeline is a dark solid arrow, events are short vertical arrows, ticks aligned |
| UI mockup | Window outer frame square; title bar dark, menu/status bars soft, buttons left-aligned, tabs in an equal-width grid, tree indentation 5 mm |

## 10. File Structure and Technical Rules

### 10.1 Visio Objects in a docx
- Each object is a `<w:object>` in `word/document.xml`, containing:
  - `<v:shape style="width:…pt;height:…pt">` display size, with `<v:imagedata r:id>` pointing at the preview image (usually EMF in the source manuscript);
  - `<o:OLEObject ProgID="Visio.Drawing.15" r:id>` pointing at `word/embeddings/Microsoft_Visio_DrawingN.vsdx`;
  - `w:dxaOrig / w:dyaOrig` original size, in twips (mm/25.4×1440).
- The rId-to-file mapping is in `word/_rels/document.xml.rels`. The order of `<w:object>` elements in document.xml is the document order; the embedding file numbers (Drawing, Drawing1…) do **not** necessarily equal the document order — always cross-check against the actual rId.

### 10.2 Modifying One Figure Means Doing All of These
1. Change `visio/pages/pageN.xml` inside the vsdx (shapes, connectors, Connects).
2. Change `PageWidth` / `PageHeight` in `visio/pages/pages.xml` (in inches).
3. Generate a new preview PNG (≥ 250 dpi), put it in `word/media/`, change the corresponding imagedata relationship's Target to it; delete old preview files no longer referenced.
4. Change the `<v:shape>` `style` width/height (pt) and `w:dxaOrig/w:dyaOrig` (twips), and change `<o:lock aspectratio="f"/>` inside `<v:shape>` to `"t"`.
5. Make sure `[Content_Types].xml` has `<Default Extension="png" ContentType="image/png"/>`.
6. When repackaging the vsdx and docx, put `[Content_Types].xml` as the first entry and keep the rest as close to the original order as possible.
7. Keep the vsdx's original document.xml, stylesheets, master shapes, and other parts; replace only the page content, to avoid errors when Visio opens it.

### 10.3 Visio Page XML Essentials
- Namespace `http://schemas.microsoft.com/office/visio/2012/main`; units in inches; **Y axis pointing up** (when laying out top-to-bottom, y_visio = page height − y).
- 2-D shapes: `Type="Shape"`, style references `LineStyle/FillStyle/TextStyle="0"`, all formatting written as local cells (FillForegnd, LineColor, LinePattern, LineWeight, Rounding, etc.).
- Order: Cells first, Sections (Character, Paragraph, Geometry) after, `<Text>` last.
- Character rows carry `Font`, `AsianFont`, `Size` (inches, pt/72, `U="PT"`), `Color`, `Style` (1 = bold), `LangID`.
- Connectors: 1-D shapes, written with `BeginX/BeginY/EndX/EndY`, `ObjType=2`, `EndArrow=13`, and polyline Geometry; endpoint gluing, size formulas, and connector text per §7.2 and §7.3.
- When there is only `<Connects>` and the endpoint cells have no `PAR(PNT(...))` formula, dragging a shape in Visio won't move the endpoints; this is the most common "looks connected but isn't actually glued" case.
- If a container's title band is a standalone shape, it won't follow when the container is dragged; to move them together, write the title into the container's own text, or group them in Visio.
- In an environment without Visio, drag testing isn't possible; with Windows + Visio, use `visio_finish.py --verify-only` for automatic drag verification; otherwise ask the user to double-click the figure and drag one shape to confirm.

## 11. Toolchain (Tools Actually Used with This Guide)

### 11.1 Agent Use
Any LLM agent that can run a shell, read/write files, and view PNGs as images can work from this guide; look at the preview PNG after finishing each figure.

### 11.2 Runtime Environment and Libraries
- Python 3 standard library: `zipfile` (unpacking/packaging docx/vsdx), `re` (locating `<w:object>`, changing size attributes), `xml.etree.ElementTree` / `lxml` (parsing and validating XML), `xml.sax.saxutils.escape` (escaping shape text), `math` (connector angles and local coordinates).
- Pillow (PIL): renders preview PNGs (`ImageDraw.rounded_rectangle`, `text(anchor='mm')`, hand-drawn dashes).
- Fonts: defaults by operating system (`scripts/font_support.py`: Windows Arial + Microsoft YaHei, macOS Helvetica Neue + PingFang, Linux fontconfig detection); PNG previews and Visio XML use the same font families; override with `VISIO_FONT_LATIN` / `VISIO_FONT_CJK` and `VISIO_FONT_*_FILE` (old variables `VISIO_KIT_FONT_R/B` compatible). If no font is found, it errors out; no silent substitution.
- It can be done with no Visio / LibreOffice: previews are drawn by the script itself and don't depend on Office rendering the vsdx.

### 11.3 XML Backend Scripts (scripts/xml_backend/)
| File | What it does | Usage |
|---|---|---|
| `scripts/xml_backend/extract_vsdx.py` | Extracts all embedded vsdx files from the source docx as that document's **template directory** (keeping each figure's original stylesheets, windows, and other parts), and prints the body-order → embedding-suffix cross-reference | `python3 extract_vsdx.py 原稿.docx work/templates` |
| `assets/templates/blank_vsdx/` | Blank vsdx template (single page, no shapes, stylesheet includes the 0th style used by vlib, author info and thumbnails removed) | Used automatically for new figures or when the source has no corresponding embedding |
| `scripts/xml_backend/inspect_visio_docx.py` | **General diagnostics**: per object, lists page size, Word size, scaling, non-uniform stretching, aspect-ratio lock, formula errors, diagonal lines, unglued connectors (counted by Connects), minimum effective font size | `python3 inspect_visio_docx.py file.docx` |
| `scripts/xml_backend/vlib.py` | **General drawing DSL**: style palette, containers, orthogonal connectors → Visio page XML + PNG preview; automatically adds connection points to shapes, glues connector endpoints that land on borders, writes connector text into the connector itself | See §12 |
| `scripts/xml_backend/vsdxpkg.py` | vsdx packaging library: template-directory copy + new page1.xml + page resized to the figure → vsdx | Called by the two scripts below |
| `scripts/xml_backend/build_docx.py` | Writes back the docx: replaces embedded vsdx files, preview PNGs, Word object sizes, and aspect-ratio locking, figure by figure; Content_Types | `python3 build_docx.py 原稿.docx figspecs.py work/templates 输出.docx --preview work/png` |
| `scripts/xml_backend/build_vsdx.py` | Generates a standalone .vsdx + PNG preview directly, without going through a docx | `python3 build_vsdx.py figspecs.py out/ [图号...]` |
| `scripts/xml_backend/visio_finish.py` | **Windows + local Visio verification and finishing (optional)**: cached-value checks, gluing checks, connector geometry first/last-point checks, automatic drag tests; forced recalculation, EMF export (keeps the PNG when the aspect ratio doesn't match), Visio save-as of the vsdx, write-back to the docx. A single figure failing doesn't affect the rest; exit code 1 on problems | `python visio_finish.py 生成.docx 终稿.docx --report r.json`; add `--verify-only` for verification only |
| `assets/example_xml_flowchart.py` | Minimal example: a flowchart with a decision branch | `python scripts/xml_backend/build_vsdx.py assets/example_xml_flowchart.py out/` |

### 11.4 Division of Labor
- Generation always uses the XML backend; it runs on any platform, with deterministic, comparable results.
- With Windows + Visio, run `visio_finish.py` before delivery: first `--verify-only` to read the report, then the formal finishing pass if there are no issues. Without Visio, deliver the build_docx output directly (PNG previews).
- Only complex source figures with master shapes, themes, or a need for local edits are modified directly with COM on the original figure.

## 12. Drawing DSL Usage (`vlib.py`)

Coordinate units mm, origin top-left, y downward; converted automatically to Visio inches, Y upward, on export.

```python
from vlib import Fig, LINE, GRAY, BLUE
f = Fig(170, 110)                                   # Page width 170 mm, page height 110 mm
f.container('sys', 4, 4, 162, 60, 'This system', tw=30)  # Container + top-left title band (dashed=True for an external container)
f.box('a', 10, 18, 30, 12, 'Import configuration')  # Style defaults to node
f.box('b', 50, 18, 40, 12, 'Processing core', 'key')  # Styles: node/key/dark/ext/soft/ok/warn
f.conn('a', 'r', 'b', 'l')                          # From a's right midpoint to b's left midpoint, orthogonal automatically
f.conn('b', 'r', 'c', 'l', label='Pass')            # Decision branch: label written into the connector's own text, placed at the longest segment's midpoint automatically
f.conn('c', 'l', 'a', 'b', via=[(20, 70)], dashed=True)  # via gives bend points manually
f.line([(x1, y1), (x2, y2)], src='a', dst='b')      # Fully manual polyline; src/dst optional, endpoints landing on borders glue automatically
f.render_png('fig.png', S=10)                       # S=pixels/mm, 10 ≈ 254 dpi
print(f.glue_report())                              # (glued endpoint count, total endpoint count); check that unglued ones are all axes/ticks
xml = f.visio_xml()                                 # Written into the vsdx as visio/pages/page1.xml
```

- `conn` edges: `t/b/l/r`; `fa/fb` are proportional positions along an edge (0–1), for drawing multiple lines from the same edge.
- Manual polyline endpoints must land exactly on shape borders; branch lines start from the source shape's edge, not from the middle of another line.
- `legend()` and `text(..., 'note')` remain in the library, but are not used by default per §6 and §8.
- Helper-function ideas: `row(n, x0, x1, gap)` computes equal-width horizontal rows; `chain(f, ids)` connects them left-to-right in sequence.
- Fonts are taken from the operating-system default; override with `VISIO_FONT_LATIN` / `VISIO_FONT_CJK` (or the old variables `VISIO_KIT_FONT_R/B`); no code changes needed.

## 13. Pre-Delivery Self-Check Checklist

- [ ] Page width = text-block width, 100% in Word with aspect ratio locked (diagnostic script shows no "scaling/stretching/unlocked")
- [ ] Fonts unified: body 9 pt, connector text 8 pt, effective font size ≥ 7.5 pt, no formula errors
- [ ] Only one palette used; dark ≤ 2
- [ ] Rectangles square (except terminators); no legends, no on-figure annotations
- [ ] All connectors orthogonal, not crossing unrelated shapes
- [ ] All connectors with shapes at both ends glued: BeginX/EndX have `PAR(PNT(...))` formulas, `<Connects>` one-to-one; the only unglued ones left are axes, ticks, separators
- [ ] No branches starting from the middle of a line
- [ ] Connector text only on decision branches, written in the connector's own text, with no standalone label text boxes
- [ ] Text not overflowing, hierarchy indentation tidy
- [ ] Consistent with body text and figure captions; no numbering conflicts
- [ ] Preview images updated in sync, old previews deleted
- [ ] All XML parses; docx readable as a zip
- [ ] Every preview visually inspected with an image viewer
- [ ] Delivery notes ask the user to drag a shape in Visio to verify the follow behavior
- [ ] Modification records written in the project ledger

## 14. Project Ledger (Recommended Approach)

Keep a separate ledger file for each modified document (e.g. `PROJECT_FIGURES.md`), recording:
- Figure number → embedding file cross-reference (document order);
- The layout template used for each figure;
- Deviations from this guide's defaults (e.g. page width 150 mm, a different palette) and why;
- Modification records: date, version, per-figure changes, reasons for content changes.

This guide holds only general rules; when new general experience is found (a new problem type, a new template, a new tool usage), update this guide and record it below.

## 15. Guide Change Log

- v2.3 (2026-11-19): submitted together with the XML backend; rules drawn from the unified redraw and review revisions of a 17-figure Chinese design proposal.
