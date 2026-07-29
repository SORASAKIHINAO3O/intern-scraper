---
name: wps-doc
zh_name: "WPS 文档助手"
en_name: "WPS Doc"
emoji: "📄"
description: 用 python-docx 生成、编辑、读取 WPS 文档(.docx)。当用户要"写周报/写合同/写报告/生成 Word 文档/整理成 docx/读取 docx 内容/给文档加表格"等中文办公文档需求时使用。生成的是标准 .docx,WPS Writer 与 Microsoft Word 通用。
category: document
tags: ["wps", "word", "docx", "office", "report", "中文办公"]
---

# WPS 文档助手

让 Codex 直接帮用户生成 / 编辑 / 读取 **WPS 文档(.docx)**。WPS 与 Office 同格式,生成的 .docx 两边都能直接打开。

## 何时用
用户说到"写周报 / 合同 / 报告 / 通知 / 说明书""生成 Word / docx 文档""把内容整理成文档""读一下这个 docx""给文档加个表格"等。

## 依赖
```bash
pip install python-docx
```
(无需安装或打开 WPS / Word 本身,纯文件生成。)

## 用法:优先复用本技能自带的 `scripts/docx_tools.py`
里面已封装好常用函数,直接 import 即可:

```python
import sys
sys.path.append("scripts")  # 指向本技能的 scripts 目录
import docx_tools as dt

doc = dt.new_document()
dt.add_title(doc, "周工作报告", subtitle="研发部 · 2026 第 23 周")
dt.add_section(doc, "一、本周完成", ["完成了 A、B、C。", "推进了 D。"])
dt.add_section(doc, "二、数据概览")                      # 只加小标题
dt.add_table(doc, ["指标", "本周", "环比"],
             [["新增用户", "1,280", "+12%"], ["活跃", "5 天", "持平"]])
dt.add_section(doc, "三、下周计划", ["…"])
doc.save("周报.docx")
```

可用函数:`new_document()`、`add_title(doc, text, subtitle=None)`、`add_section(doc, heading, body=None, level=1)`、`add_table(doc, header, rows, style=...)`、`read_document(path)`。

直接跑 `python scripts/docx_tools.py 文件名.docx` 会生成一份示例周报,可用来快速验证环境。

## 直接用 python-docx 的常见配方
若需求超出封装范围,直接写 python-docx:
```python
from docx import Document
from docx.shared import Pt, RGBColor
doc = Document()
doc.add_heading("标题", level=1)
p = doc.add_paragraph("正文,")
p.add_run("加粗部分").bold = True
doc.add_picture("图.png")          # 插图
doc.save("out.docx")
```

## 读取已有文档
```python
import docx_tools as dt
print(dt.read_document("已有文档.docx"))   # 返回全文文本
```

## 导出 PDF(可选,需本机装了 WPS 或 Office)
python-docx 本身不导 PDF。若用户要 PDF,在 Windows 上可用 `docx2pdf`(底层调 WPS/Word):
```bash
pip install docx2pdf
python -c "from docx2pdf import convert; convert('周报.docx', '周报.pdf')"
```
没装 WPS/Office 的环境就只交付 .docx。

## 要点
- 默认简体中文,不要无谓中英混排。
- 文件名带中文没问题;保存路径用用户指定目录。
- 生成后口头确认产物路径,必要时用 `read_document` 抽查前几行。
