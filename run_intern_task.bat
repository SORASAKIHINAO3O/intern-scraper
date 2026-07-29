@echo off
chcp 65001 >nul
cd /d "C:\Users\sunday\AppData\Roaming\reasonix\global-workspace"

set OUTPUT_DIR=C:\Users\sunday\Desktop\intern_reports
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

set DATETAG=%date:~0,4%%date:~5,2%%date:~8,2%
set REPORT_FILE="%OUTPUT_DIR%\intern_report_%DATETAG%.md"

set PYTHONIOENCODING=utf-8
python intern_scraper.py --output %REPORT_FILE%

echo 报告已保存至: %REPORT_FILE%
echo 完成时间：%date% %time%
