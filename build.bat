@echo off
chcp 65001 >nul
title 实习手册封面生成器 - 打包脚本
cd /d "%~dp0"

echo ============================================================
echo    实习手册封面生成器   打包脚本
echo ============================================================
echo.

rem ---------- 1. 检查 Python ----------
set PY=
where python >nul 2>nul && set PY=python
if not defined PY (
    where py >nul 2>nul && set PY=py
)
if not defined PY (
    echo [错误] 没有找到 Python。
    echo.
    echo 请先安装 Python 3.9 或更高版本：
    echo     https://www.python.org/downloads/
    echo.
    echo 安装时一定要勾选 "Add Python to PATH"
    echo.
    pause
    exit /b 1
)
echo [1/4] 使用 Python: %PY%
%PY% --version
echo.

rem ---------- 2. 询问打包方式 ----------
echo [2/4] 选择打包方式
echo.
echo   1 = 单文件 exe  (只有一个文件，方便拷贝，推荐)
echo   2 = 文件夹版    (启动更快，但会带一个文件夹)
echo.
set MODE=1
set /p MODE="请选择 [1]: "
if "%MODE%"=="2" (set MODE=2) else (set MODE=1)
echo.

rem ---------- 3. 安装依赖 ----------
echo [3/4] 安装依赖库（首次会慢一些，请耐心等待）...
%PY% -m pip install --upgrade pip --quiet
%PY% -m pip install --quiet pillow python-docx pywin32 pyinstaller
if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败，请检查网络连接后重试。
    pause
    exit /b 1
)
echo       依赖安装完成
echo.

rem ---------- 4. 打包 ----------
echo [4/4] 正在打包，请稍候...
echo.

set NAME=实习手册封面生成器
rem 注意：这里刻意不写 --hidden-import=PIL.ImageTk
rem Pillow 10+ 的 hook 会排除 tkinter，导致 ImageTk 漏打包、预览变空白
rem 程序内部已改用 tkinter 原生 PhotoImage，无需 ImageTk
set HIDDEN=--hidden-import=win32print --hidden-import=win32ui --hidden-import=win32con --hidden-import=win32com.client --hidden-import=PIL.ImageWin --hidden-import=docx --hidden-import=docx.oxml --hidden-import=docx.opc.constants

if "%MODE%"=="1" (
    %PY% -m PyInstaller --noconfirm --clean --onefile --windowed ^
        --name "%NAME%" %HIDDEN% ^
        --distpath ".\dist" --workpath ".\build" --specpath ".\build" ^
        cover_maker.py
) else (
    %PY% -m PyInstaller --noconfirm --clean --onedir --windowed ^
        --name "%NAME%" %HIDDEN% ^
        --distpath ".\dist" --workpath ".\build" --specpath ".\build" ^
        cover_maker.py
)

if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请查看上方错误信息。
    pause
    exit /b 1
)

rem ---------- 5. 自检 ----------
echo.
echo [自检] 正在验证产物能否正常启动...
"%~dp0dist\%NAME%.exe" --help >nul 2>&1
if errorlevel 1 (
    echo   [警告] 自检未通过！双击 exe 可能报错，请把上面的日志发出来排查。
) else (
    echo   [自检] 通过，产物运行正常。
)

echo.
echo ============================================================
echo   打包完成！
echo.
echo   成品位置:  %cd%\dist
echo.
echo   把 dist 里的整个文件夹拷到任意 Windows 电脑上，
echo   双击 "%NAME%.exe" 即可使用，无需安装 Python。
echo ============================================================
echo.
pause
