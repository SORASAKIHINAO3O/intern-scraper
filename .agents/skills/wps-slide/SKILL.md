---
name: wps-slide
zh_name: "WPS 演示助手"
en_name: "WPS Slide"
emoji: "📽️"
description: 用 python-pptx 生成 WPS 演示(.pptx)。当用户要"做 PPT/做幻灯片/产品介绍/方案汇报/路演/演示文稿"等需求时使用。生成的是标准 .pptx,WPS 演示与 Microsoft PowerPoint 通用。
category: slides
tags: ["wps", "powerpoint", "pptx", "office", "deck", "中文办公"]
---

# WPS 演示助手

让 Codex 直接帮用户生成 **WPS 演示(.pptx)**。WPS 与 Office 同格式,生成的 .pptx 两边都能直接打开。

## 何时用
用户说到"做个 PPT / 幻灯片""产品介绍 / 方案汇报 / 项目路演 / 工作汇报演示""把这些内容做成演示"等。

## 依赖
```bash
pip install python-pptx
```
(纯文件生成,无需安装或打开 WPS / PowerPoint。)

## 用法:优先复用本技能自带的 `scripts/pptx_tools.py`
```python
import sys
sys.path.append("scripts")
import pptx_tools as pt

prs = pt.new_presentation()                       # 16:9
pt.add_title_slide(prs, "产品介绍", "副标题")
pt.add_bullets_slide(prs, "目录", ["背景", "方案", "案例", "下一步"])
pt.add_section_slide(prs, "解决方案")             # 章节分隔页(珊瑚色大字)
pt.add_bullets_slide(prs, "三大能力", [
    "能力一", ("子要点缩进一级", 1),
    "能力二",
])
prs.save("产品介绍.pptx")
```
可用函数:`new_presentation()`、`add_title_slide(prs, title, subtitle="")`、`add_bullets_slide(prs, title, bullets)`(bullets 项可为 `"文本"` 或 `("文本", 缩进级别)`)、`add_section_slide(prs, text)`。

直接跑 `python scripts/pptx_tools.py 文件名.pptx` 生成一份示例演示(封面/目录/章节/要点 4 页)。

## 直接用 python-pptx 的常见配方
```python
from pptx import Presentation
from pptx.util import Inches, Pt
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])   # 仅标题
slide.shapes.title.text = "标题"
box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(3))
box.text_frame.text = "正文"
slide.shapes.add_picture("图.png", Inches(1), Inches(2), width=Inches(4))
prs.save("out.pptx")
```
常用版式索引:`0` 封面、`1` 标题+内容、`5` 仅标题、`6` 空白。

## 内容建议(让 PPT 更专业)
- 每页标题写**结论句**(如"营收同比 +30%"),不要只写标签(如"营收")。
- 先封面、再目录,然后按"背景 → 问题 → 方案 → 证据 → 行动"组织。
- 一页一个核心观点,要点不超过 6 条,默认简体中文。

## 要点
- 复杂排版/配图需求,可先生成结构再逐页补内容。
- 生成后口头确认产物路径与页数。
