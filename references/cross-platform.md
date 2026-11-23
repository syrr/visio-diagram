[中文文档 (Chinese)](cross-platform.zh-CN.md)

# XML Backend: Cross-Platform Diagram Output Without Visio

`scripts/xml_backend/` generates Visio page XML directly and packages it as vsdx, with no need for Windows, Visio, or pywin32 — it runs on Linux, macOS, and Windows alike. The output is still natively editable Visio objects: connection points, endpoint gluing (`PAR(PNT())` formulas + `<Connects>`), and connector-owned text are all written per the Visio spec, so connector endpoints follow when you drag a shape.

## How It Divides Work with the COM Backend
| | COM (visio_lib.py) | XML backend |
|---|---|---|
| Runtime environment | Windows + desktop Visio | Any platform, only Python 3.8+ and Pillow needed |
| Speed and determinism | Slower, affected by Visio version | Seconds, same input gives same output, diffable |
| Word preview | Vector EMF | High-DPI PNG (automatically regenerated after double-click editing in Visio) |
| Formula cached values | Recalculated by Visio on the spot | Computed and written by the script itself |
| Best for | Needing EMF, or making local edits to an existing diagram | Batch generation, whole-page redraws, CI, AI agent sandboxes |

The two can be combined: generate with the XML backend, then use `visio_finish.py` on a Windows machine with Visio for verification and finishing (recalculation, EMF export, save-as from Visio). If that step fails, delivery is unaffected.

## Usage
```bash
pip install pillow
# 1. Standalone vsdx + PNG preview
python scripts/xml_backend/build_vsdx.py assets/example_xml_flowchart.py out/

# 2. Redraw figures already embedded in a docx (source manuscript is read-only)
python scripts/xml_backend/inspect_visio_docx.py 方案.docx              # Diagnose the current state
python scripts/xml_backend/extract_vsdx.py 方案.docx work/templates     # Extract per-figure templates, printing "order → embedding suffix"
#    Write work/figspecs.py: FIGS = [(figure number, embedding suffix, generator function), ...]
python scripts/xml_backend/build_docx.py 方案.docx work/figspecs.py work/templates out.docx --preview work/png
python scripts/xml_backend/inspect_visio_docx.py out.docx               # Re-check

# 3. Tests (no Visio needed)
python -m unittest discover -s tests -v

# 4. Optional: Windows + Visio verification and finishing (requires pywin32)
python scripts/xml_backend/visio_finish.py out.docx --verify-only --report r.json
python scripts/xml_backend/visio_finish.py out.docx final.docx
```

## Themes
A theme changes only colors, never semantic roles, layout, font sizes, or connector rules; themes are defined in `XML_PALETTES` in `scripts/theme.py`, sharing a registry (`THEMES` in `theme.py`) with the COM themes, so the same name means the same thing on both routes. The default is `tech-cn`; omitting the theme leaves the appearance unchanged.

| Theme | Suited for | Primary color |
|---|---|---|
| `tech-cn` (default) | Chinese engineering documents | Dark blue #1F3864 |
| `mono` | Strict black-and-white printing | All grayscale; ok/warn distinguished only by fill darkness |

```python
f = Fig(150, 70, theme="mono")   # Single figure (black-and-white printing)
```
```bash
python scripts/xml_backend/build_docx.py 方案.docx work/figspecs.py work/templates out.docx --theme mono
VISIO_THEME=mono python scripts/xml_backend/build_vsdx.py assets/example_xml_flowchart.py out/   # Environment variable also works
```
Use only one theme per document (see `visio-style-rules.md` §6). ok/warn are semantic colors, identical across themes except mono; every theme's text/fill contrast is ≥ 4.5:1 (see the contrast tests in `tests/test_xml_backend.py`).

## Files
- `vlib.py`: drawing DSL (box / text / container / line / conn / legend), outputs Visio page XML and PNG previews; units mm, origin at top-left, Y downward.
- `vsdxpkg.py`: template directory + new page1.xml → vsdx, resizing the page to the figure dimensions.
- `extract_vsdx.py`: extracts all embedded vsdx files from a docx, preserving each figure's original stylesheets and window parts.
- `build_docx.py`: replaces embedded vsdx files and previews figure by figure, sets object sizes, and locks the aspect ratio.
- `build_vsdx.py`: produces a standalone vsdx directly, without going through a docx.
- `inspect_visio_docx.py`: read-only diagnostics (sizes, scaling, stretching, font sizes, unglued connectors, etc.).
- `visio_finish.py`: Windows + Visio checks of cached values, gluing, connector geometry first/last points, and automatic drag tests (checking both endpoints and geometry after dragging), plus recalculation, EMF export, save-as, and write-back.
- `assets/templates/blank_vsdx/`: blank template (single page, no shapes, author info and thumbnails removed); used automatically when the source manuscript has no corresponding embedding. Only single-page templates are supported; a multi-page template errors out directly.
- `tests/test_xml_backend.py`: smoke and regression tests (gluing, zero-width/zero-height connector geometry formulas, title bands excluded from gluing, alignment, error messages, synthetic docx round-trips, template directory cleanup).

## Fonts
The default font is chosen by operating system (`scripts/font_support.py`, shared by the COM and XML routes): Windows = Arial + Microsoft YaHei; macOS = Helvetica Neue + PingFang SC; Linux = system fonts detected via fontconfig (e.g. Noto Sans CJK). PNG previews use the same font; in Visio XML, `Font`/`AsianFont` carry the Latin/CJK font families respectively. Override with environment variables: `VISIO_FONT_LATIN` / `VISIO_FONT_CJK` (font families) and `VISIO_FONT_{LATIN,CJK}{,_BOLD}_FILE` (preview metric font files); the old variables `VISIO_KIT_FONT_R/B` (.ttc index `VISIO_KIT_FONT_INDEX`) remain compatible. If no font is found, it errors out rather than silently substituting an unrelated font.

## Known Limitations
- Whole-page redraws only; no local edits to an existing diagram, and no handling of master shapes, theme inheritance, or stencils.
- Connectors are 1-D shapes with explicit geometry rows, every geometry point written as a formula: when dragged, endpoints and the line body follow together, and middle bend points scale proportionally but do not re-route automatically (same as the COM backend's connect()).
- Cached values for complex formulas are computed by the script; when possible, verify them with `visio_finish.py --verify-only`.
- `visio_finish.py` has been regression-tested against real Visio (2026-11-19, Visio Pro + Win10): gluing checks / drag tests / cached values / EMF export / write-back all passed end to end; EMF export uses an "opaque white page frame underneath" so the figure extent equals the page (a transparent frame does not expand the export extent).

For the complete style rules and review checklist, see `visio-style-rules.md`.
