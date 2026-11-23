[中文文档 (Chinese)](cells-reference.zh-CN.md)

# Visio CellsU Quick Reference and Coordinate Conversion

## 1. Coordinate Conversion (Where Mistakes Happen Most)

Inside Visio: **the origin is at the bottom-left corner of the page, Y axis pointing up**, default unit inches.
This skill's `visio_lib` exposes a uniform convention of **top-left origin, Y downward, mm**; the conversion is just two lines:

```python
PinX = x + w / 2.0                    # Horizontal direction is the same in both
PinY = pageH - ytop - h / 2.0         # Vertical direction must be flipped
Width  = w
Height = h
```

Lines (`DrawLine`) work the same way, flipping only Y:

```python
BeginY = pageH - y1
EndY   = pageH - y2
```

**Always include a unit when writing a Cell**, otherwise Visio interprets it as inches:

```python
s.CellsU("PinX").FormulaU = "25 mm"      # ✓
s.CellsU("PinX").FormulaU = "25"         # ✗ Will be read as 25 inches
s.CellsU("Char.Size").FormulaU = "8.5 pt"   # Font size in pt
```

`CellsU` values are always **strings**; booleans use `"0"` / `"1"` — do not pass Python's `True`.

---

## 2. PageSheet (Page)

| Cell | Notes | Example |
|---|---|---|
| `PageWidth` / `PageHeight` | Page size | `"152 mm"` |
| `LeftMargin` / `RightMargin` / `TopMargin` / `BottomMargin` | Page margins; **set to 0** so content can be laid out freely | `"0 mm"` |

```python
pg = doc.Pages.Item(1)
pg.PageSheet.CellsU("PageWidth").FormulaU = "152 mm"
```

---

## 3. Shape

### Geometry
| Cell | Notes |
|---|---|
| `PinX` / `PinY` | Shape **center** coordinates |
| `Width` / `Height` | Size |
| `BeginX/BeginY/EndX/EndY` | Line-type shapes only (`DrawLine`) |
| `Rounding` | Corner radius, `"0 mm"` square, `"1.5 mm"` rounded |

### Fill
| Cell | Notes |
|---|---|
| `FillPattern` | `0` no fill, `1` solid, 2–24 various patterns |
| `FillForegnd` | Fill color, `"RGB(222,235,247)"` |

### Lines
| Cell | Notes |
|---|---|
| `LinePattern` | `0` no line, `1` solid, `2` dashed, `4` dash-dot (dozens of codes in total, see below) |
| `LineColor` | `"RGB(0,0,0)"` |
| `LineWeight` | `"0.75 pt"` |
| `EndArrow` / `BeginArrow` | Arrow style codes |
| `Shadows` | Usually set to `"0"` to turn shadows off |

### Text
| Cell | Notes |
|---|---|
| `Char.Font` | Font name, **must be quoted**: `'"微软雅黑"'` |
| `Char.Color` | `"RGB(38,38,38)"` |
| `Char.Size` | `"8.5 pt"` |
| `Para.HorzAlign` | `0` left, `1` center, `2` right, `3` justified |
| `VerticalAlign` | `0` top, `1` middle, `2` bottom |
| `LeftMargin` / `RightMargin` | Horizontal inner padding between text and box edge |
| `TopMargin` / `BottomMargin` | Vertical inner padding |

Text content is assigned as a property, not a Cell:

```python
s.Text = "标题\n说明文字"        # Line breaks use \n
```

> Forgetting the quotes on `Char.Font` is the most common silent failure: no error, but the font doesn't take effect.

---

## 4. Code Values (Trust the Actual Rendering)

Visio's line-pattern and arrow codes number in the dozens and vary slightly by version. **Reliable approach**: set it once manually in the Visio UI, then read back `CellsU(...).FormulaU` to see the code.

Common values verified by this skill:

| Use | Value |
|---|---|
| No arrow | `EndArrow = 0` |
| Solid triangular arrow | `EndArrow = 4` |
| Hollow arrow | `EndArrow = 1` |
| Solid line | `LinePattern = 1` |
| Dashed line | `LinePattern = 2` |
| Dash-dot line | `LinePattern = 4` |

---

## 5. The Fixed Pattern for Creating Shapes

Visio COM has no direct API for "create a rectangle from a rectangle". The pattern is **create a default shape first, then position it**:

```python
s = pg.DrawRectangle(0, 0, 1, 1)     # First create a 1x1 one
s.CellsU("PinX").FormulaU = "..."    # Then change size and position
s.CellsU("PinY").FormulaU = "..."
s.CellsU("Width").FormulaU  = "60 mm"
s.CellsU("Height").FormulaU = "20 mm"
```

Use `pg.DrawLine(0, 0, 1, 1)` for lines, then change `BeginX/BeginY/EndX/EndY`.

---

## 6. Export and Save

```python
doc.SaveAs(r"D:\out\图A.vsdx")     # Native format, still editable
pg.Export(r"D:\out\图A.emf")       # Vector, for embedding in Word/PPT
pg.Export(r"D:\out\图A.png")       # Bitmap preview
```

- `Page.Export` **crops to the content bounds** (excluding page margins), 96 DPI — so no manual trimming of white space is needed.
- The extension determines the format, handled by Visio's export filters.
- `Export` exports only the current page; for multi-page documents, call it once per page.

---

## 7. COM Lifecycle

```python
app = win32.Dispatch("Visio.Application")
app.AlertResponse = 7      # Suppress popups (otherwise it may hang waiting for a manual click)
app.Visible = 0
doc = app.Documents.Add("")
...
doc.Close()
app.Quit()
```

**Always call `Quit()`**, otherwise VISIO.EXE stays behind as a background process and accumulates over repeated runs.
