@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   人员核验 PDF 自动校验
echo ========================================
echo.

REM 检查虚拟环境
if not exist ".venv\Scripts\python.exe" (
    echo [错误] 未找到虚拟环境 .venv
    echo 请先在项目目录执行:
    echo   python -m venv .venv
    echo   .venv\Scripts\pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

REM 确保 inputfile 目录存在
if not exist "inputfile" mkdir inputfile

REM 统计 PDF 数量（Windows 不区分大小写，*.pdf 即可匹配 .PDF）
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

REM 运行核验
".venv\Scripts\python.exe" main.py

echo.
echo ========================================
echo   核验完成，报告已输出到 output 目录
echo ========================================
echo.
pause
