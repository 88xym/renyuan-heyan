@echo off
rem ========================================
rem   Person Check - PDF Auto Verify
rem   NOTE: this bat must be saved in ANSI/GBK
rem ========================================
cd /d "%~dp0"

echo ========================================
echo   人员核验 PDF 自动校验
echo ========================================
echo.

rem check virtual env
if not exist ".venv\Scripts\python.exe" (
    echo [错误] 未找到虚拟环境 .venv
    echo 请先在项目目录执行:
    echo   python -m venv .venv
    echo   .venv\Scripts\pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

rem ensure inputfile exists
if not exist "inputfile" mkdir inputfile

rem count pdf files
set count=0
for %%f in (inputfile\*.pdf) do set /a count+=1

if %count%==0 (
    echo [提示] inputfile 目录下没有找到 PDF 文件。
    echo 请把待核验的 PDF 放入 inputfile 目录后再运行。
    echo.
    pause
    exit /b 0
)

echo 找到 %count% 个 PDF 文件，开始核验...
echo.

rem run verify
set PYTHONIOENCODING=gbk
".venv\Scripts\python.exe" main.py

rem auto verify -> PDF report (per-person one page)
echo.
echo [PDF] 正在生成每人一页核验报告（自动读取 OCR 缓存）...
echo [PDF] 如扫描件未 OCR，会自动先 OCR，请耐心等待...
".venv\Scripts\python.exe" tools\auto_verify.py

echo.
echo ========================================
echo   核验完成，报告已输出到 output 目录
echo ========================================
echo.
pause
