---
name: intern-scraper
description: 抓取字节/腾讯/阿里/百度/美团/快手/拼多多/小红书的2027届实习岗位，输出表格
---

# intern-scraper — 互联网大厂实习岗位信息爬虫

使用 Playwright + Edge 浏览器自动抓取8家互联网大厂（字节跳动、腾讯、阿里巴巴、百度、美团、快手、拼多多、小红书）的2027届实习/校招岗位信息。

## 用法

```bash
cd "C:\Users\sunday\AppData\Roaming\reasonix\global-workspace"
set PYTHONIOENCODING=utf-8
python intern_scraper.py                     # 输出到控制台
python intern_scraper.py --output report.md   # 保存到文件
python intern_scraper.py --company 字节跳动,小红书  # 只抓取指定公司
```

## 输出格式

Markdown 表格，包含：公司名称、岗位/项目名称、工作地点、岗位要求、薪资范围、备注链接。

## 定时任务

已配置 Windows 计划任务 `InternScraperTask`，每天 09:00 自动运行，报告保存至桌面 `intern_reports` 目录。

## 如果需要手动运行

直接执行 `run_intern_task.bat` 即可。

## 脚本位置

- 脚本: `C:\Users\sunday\AppData\Roaming\reasonix\global-workspace\intern_scraper.py`
- 批处理: `C:\Users\sunday\AppData\Roaming\reasonix\global-workspace\run_intern_task.bat`
- 输出目录: `C:\Users\sunday\Desktop\intern_reports\`
