# visio-diagram

[English](README.md)

用代码画可编辑的 Visio 图，并嵌进 Word —— 不用再手动拖形状。

每张图导出三件套：`.vsdx`（原生可编辑）· `.emf`（矢量，供 Word / PPT）· `.png`（预览）。

## 为什么用它

- **图一直可编辑：** 连接线胶在连接点上，挪框时端点跟随，不用再手动重画箭头。
- **自动自检：** 出图前做文字溢出自检和箭头对齐审计，提前发现版式问题。
- **Word 就绪：** 按占位符嵌入，或嵌成双击可进 Visio 改图的 OLE 对象。
- **不装 Visio 也能用：** XML 后端在 Linux / macOS / Windows 上生成同样的可编辑 `.vsdx`，只需 Python 和 Pillow。

## 快速安装

```bash
git clone https://github.com/syrr/visio-diagram.git
cd visio-diagram
pip install pywin32 pillow   # 只有 Windows + Visio 才需要 pywin32
```

没装 Visio？只需 `pip install pillow`，XML 后端不需要别的。

## 快速上手

有 Visio（Windows）：

```python
import sys
sys.path.insert(0, "scripts")
from visio_lib import VDraw

with VDraw(150, 108) as d:      # 画布 mm，左上原点
    d.lane(0, 0, 150, 74, "处理链路")
    d.step_box(6, 12, 42, 26, "① 环节一", "说明文字")
    d.h_arrow(48, 25, 60)
    d.check()                   # 文字溢出自检
    d.export_all("out/图N1")    # 导出 .vsdx / .emf / .png
```

把图嵌进 Word：

```bash
python scripts/word_insert.py 方案.docx --placeholder "{{FIG:N2}}" --image 图N2.emf
```

没装 Visio（任意平台）：

```bash
python scripts/xml_backend/build_vsdx.py assets/example_xml_flowchart.py out/
```

## 作为 LLM agent 技能

克隆到 agent 的技能目录，然后用一句话要图就行：

```bash
git clone https://github.com/syrr/visio-diagram.git ~/.claude/skills/visio-diagram
```

```
用 Visio 画一张三泳道的系统架构图，按文档栏宽 1:1 出图，自检后嵌到方案 docx 里
```

剩下的 agent 自己做：环境检查、量栏宽、写绘图脚本、溢出自检、对齐审计、出图并嵌入 Word。

## 了解更多

完整参考见 [SKILL.zh-CN.md](SKILL.zh-CN.md) 与 [references/](references/)，包括跨平台 XML 后端（[cross-platform.zh-CN.md](references/cross-platform.zh-CN.md)）和避坑清单（[pitfalls.zh-CN.md](references/pitfalls.zh-CN.md)）。

环境要求：COM 路线需要 Windows + 桌面版 Visio、Python 3.9–3.13、`pywin32`、`Pillow` 和中文字体（默认微软雅黑）。

## 许可证

MIT —— 见 [LICENSE](LICENSE)。
