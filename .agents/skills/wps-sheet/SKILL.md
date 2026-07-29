---
name: wps-sheet
zh_name: "WPS 表格助手"
en_name: "WPS Sheet"
emoji: "📊"
description: 用 openpyxl 生成、编辑、读取 WPS 表格(.xlsx)。当用户要"做表格/填数据/加公式/算合计/批量处理 Excel/整理成 xlsx/读取表格分析"等需求时使用。生成的是标准 .xlsx,WPS 表格与 Microsoft Excel 通用。
category: spreadsheet
tags: ["wps", "excel", "xlsx", "office", "data", "中文办公"]
---

# WPS 表格助手

让 Codex 直接帮用户生成 / 编辑 / 读取 **WPS 表格(.xlsx)**。WPS 与 Office 同格式,生成的 .xlsx 两边都能直接打开。

## 何时用
用户说到"做个表 / 填数据 / 加公式 / 算合计 / 做报表 / 台账""批量处理这些 Excel""把数据整理成表格""读一下这个 xlsx 帮我分析"等。

## 依赖
```bash
pip install openpyxl
```
(纯文件生成,无需安装或打开 WPS / Excel。)

## 用法:优先复用本技能自带的 `scripts/xlsx_tools.py`
```python
import sys
sys.path.append("scripts")
import xlsx_tools as xt

wb, ws = xt.new_workbook("一季度销售")
end = xt.write_table(ws, ["产品", "单价", "数量", "金额"],
                     [["键盘", 199, 120, None], ["鼠标", 89, 240, None]])
# 金额列用公式
xt.set_formula(ws, "D2", "=B2*C2")
xt.set_formula(ws, "D3", "=B3*C3")
xt.set_formula(ws, f"D{end+1}", f"=SUM(D2:D{end})")   # 合计
xt.autofit(ws)
wb.save("销售表.xlsx")
```
可用函数:`new_workbook(sheet_name)`、`write_table(ws, header, rows, start_row=1)`(表头自动加粗描边)、`set_formula(ws, cell, formula)`、`autofit(ws)`、`read_sheet(path, sheet=None)`。

直接跑 `python scripts/xlsx_tools.py 文件名.xlsx` 生成示例销售表(含 `=B*C` 与 `SUM` 合计)。

## 直接用 openpyxl 的常见配方
```python
from openpyxl import Workbook, load_workbook
wb = Workbook(); ws = wb.active
ws["A1"] = "姓名"; ws.append(["张三", 90])     # 逐行追加
ws["C1"] = "=AVERAGE(B2:B10)"                  # 公式
# 多工作表
ws2 = wb.create_sheet("汇总")
wb.save("out.xlsx")
```

## 批量处理一批表
```python
import glob, xlsx_tools as xt
for f in glob.glob("输入/*.xlsx"):
    rows = xt.read_sheet(f)
    # …处理 rows…
```

## 读取与分析
```python
import xlsx_tools as xt
rows = xt.read_sheet("数据.xlsx")    # 二维列表,可直接遍历/统计
```
数据量大、要做透视/统计时,也可改用 pandas(`pd.read_excel` / `to_excel`,底层同样兼容 WPS)。

## 要点
- 公式用字符串写,如 `"=B2*C2"`;WPS 打开会自动计算。
- 中文表头/中文文件名都没问题。
- 生成后口头确认路径,必要时 `read_sheet` 抽查前两行。
