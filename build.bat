@echo off
chcp 65001 >nul
echo === 视频音频提取器 — 一键打包脚本 ===
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 python，请先安装 Python 3.8+ 并勾选 "Add to PATH"
    echo        下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] 安装依赖...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
if errorlevel 1 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)

echo.
echo [2/3] 检查 ffmpeg.exe / ffprobe.exe...
set ADDBIN=
if exist ffmpeg.exe (
    echo     已找到 ffmpeg.exe，将打包进 exe
    set ADDBIN=%ADDBIN% --add-binary "ffmpeg.exe;."
) else (
    echo     [警告] 当前目录没有 ffmpeg.exe —— 剪辑拼接功能会无法使用
    echo     下载：https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip
)
if exist ffprobe.exe (
    echo     已找到 ffprobe.exe，将打包进 exe
    set ADDBIN=%ADDBIN% --add-binary "ffprobe.exe;."
) else (
    echo     [提示] 没有 ffprobe.exe —— 列表里看不到时长，但拼接仍能跑
)

echo.
echo [3/3] 打包中（首次可能较慢）...
pyinstaller --onefile --windowed --name VideoToAudio --clean %ADDBIN% main.py
if errorlevel 1 (
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo.
echo === 打包完成 ===
echo 可执行文件位于 dist\VideoToAudio.exe
echo.
pause
