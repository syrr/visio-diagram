# visio-diagram

[中文文档 (Chinese)](README.zh-CN.md)

Draw editable Visio diagrams with code, and embed them in Word — instead of dragging shapes by hand.

Each figure comes out as `.vsdx` (natively editable in Visio), `.emf` (vector, for Word / PowerPoint), and `.png` (preview).

## Why use it

- **Diagrams stay editable:** connectors are glued to connection points, so endpoints follow when you move a box. No more redrawing arrows by hand.
- **It checks itself:** text-overflow self-check and arrow-alignment audit catch layout mistakes before export.
- **Word-ready:** embed by placeholder, or as a double-click-to-edit Visio OLE object.
- **Works without Visio too:** the XML backend generates the same editable `.vsdx` on Linux / macOS / Windows, with only Python and Pillow.

## Quick install

```bash
git clone https://github.com/syrr/visio-diagram.git
cd visio-diagram
pip install pywin32 pillow   # pywin32 only needed on Windows with Visio
```

No Visio? Just `pip install pillow` — the XML backend needs nothing else.

## Quick use

With Visio (Windows):

```python
import sys
sys.path.insert(0, "scripts")
from visio_lib import VDraw

with VDraw(150, 108) as d:      # canvas in mm, top-left origin
    d.lane(0, 0, 150, 74, "Processing chain")
    d.step_box(6, 12, 42, 26, "① Step 1", "Description")
    d.h_arrow(48, 25, 60)
    d.check()                   # text-overflow self-check
    d.export_all("out/Fig N1")  # writes .vsdx / .emf / .png
```

Embed the figure in Word:

```bash
python scripts/word_insert.py proposal.docx --placeholder "{{FIG:N2}}" --image "Fig N2.emf"
```

Without Visio (any platform):

```bash
python scripts/xml_backend/build_vsdx.py assets/example_xml_flowchart.py out/
```

## Use it as an LLM agent skill

Clone it into your agent's skill directory and just ask for a diagram in plain language:

```bash
git clone https://github.com/syrr/visio-diagram.git ~/.claude/skills/visio-diagram
```

```
Use Visio to draw a three-swimlane system architecture diagram, sized 1:1 to the document column width, run the self-check, then embed it in the proposal docx
```

The agent handles the rest: environment check, column-width measurement, drawing, self-check, alignment audit, export, and Word embedding.

## Learn more

Full reference: [SKILL.md](SKILL.md) and [references/](references/) — including the cross-platform XML backend ([cross-platform.md](references/cross-platform.md)) and the pitfalls list ([pitfalls.md](references/pitfalls.md)).

Requirements: Windows + desktop Visio, Python 3.9–3.13, `pywin32`, `Pillow`, and a CJK font (default: Microsoft YaHei) for the COM route.

## License

MIT — see [LICENSE](LICENSE).
