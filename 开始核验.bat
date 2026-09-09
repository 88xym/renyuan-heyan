@echo off
rem ========================================
rem   Person Check - PDF Auto Verify
rem   NOTE: this bat must be saved in ANSI/GBK
rem ========================================
setlocal
cd /d "%~dp0"

echo ========================================
echo   人员核验 PDF 自动校验
echo ========================================
echo.

rem ========== 环境自动识别与安装 ==========
rem 1) 探测可用的 Python 3（py 启动器优先，避免 Microsoft Store 占位符）
set "PYEXE="
set "PYFULL="
py -3 --version >nul 2>nul
if not errorlevel 1 set "PYEXE=py -3"
if not defined PYEXE (
    python --version >nul 2>nul
    if not errorlevel 1 set "PYEXE=python"
)
if not defined PYEXE (
    if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set "PYFULL=%LocalAppData%\Programs\Python\Python312\python.exe"
)
if not defined PYEXE if not defined PYFULL (
    echo [环境] 未检测到 Python 3，将自动下载安装 Python 3.12（约 25MB，请联网）...
    powershell -NoProfile -Command "Invoke-WebRequest -UseBasicParsing -Uri 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe' -OutFile \"%TEMP%\python-3.12.10-amd64.exe\""
    if errorlevel 1 (
        echo [错误] Python 下载失败。请手动到 python.org 安装 Python 3.12 后重试。
        pause
        exit /b 1
    )
    echo [环境] 正在静默安装 Python 3.12...
    "%TEMP%\python-3.12.10-amd64.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_test=0 Include_doc=0 Include_tcltk=0
    if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
        set "PYFULL=%LocalAppData%\Programs\Python\Python312\python.exe"
    ) else (
        echo [错误] Python 安装失败，请手动安装 Python 3.12 后重试。
        pause
        exit /b 1
    )
    echo [环境] Python 3.12 安装完成
)

rem 2) 虚拟环境
if not exist ".venv\Scripts\python.exe" (
    echo [环境] 未找到虚拟环境，正在创建...
    if defined PYEXE (
        %PYEXE% -m venv .venv
    ) else (
        "%PYFULL%" -m venv .venv
    )
    if errorlevel 1 (
        echo [错误] 虚拟环境创建失败，请手动执行: python -m venv .venv
        pause
        exit /b 1
    )
    echo [环境] 虚拟环境创建完成
)

rem 3) 依赖检查与安装
".venv\Scripts\python.exe" -c "import pymupdf, yaml, reportlab, openpyxl, rapidocr_onnxruntime, numpy" >nul 2>nul
if errorlevel 1 (
    echo [环境] 正在安装依赖（首次需数分钟，请勿关闭窗口）...
    ".venv\Scripts\pip.exe" install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败，请检查网络后重试
        pause
        exit /b 1
    )
    echo [环境] 依赖安装完成
)

rem 4) onnxruntime DLL 检测（新电脑常缺 VC++ 运行库导致加载失败）
".venv\Scripts\python.exe" -c "import onnxruntime" >nul 2>nul
if errorlevel 1 (
    echo [环境] onnxruntime 加载失败，尝试自动安装 VC++ 运行库...
    powershell -NoProfile -Command "Invoke-WebRequest -UseBasicParsing -Uri 'https://aka.ms/vs/17/release/vc_redist.x64.exe' -OutFile \"%TEMP%\vc_redist.x64.exe\""
    if not errorlevel 1 (
        "%TEMP%\vc_redist.x64.exe" /install /quiet /norestart
    )
    echo [提示] 如仍报 DLL 错误，请手动安装微软 VC++ 运行库 x64：
    echo        https://aka.ms/vs/17/release/vc_redist.x64.exe
    echo.
)
echo [环境] 环境检查完成
echo.

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
