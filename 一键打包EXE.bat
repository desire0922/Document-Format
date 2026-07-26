@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ==========================================
echo   语文试卷一键排版 - 一键打包 EXE
echo ==========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 没有找到 Python。
    echo 请先安装 Python 3.8 或更高版本，并勾选 Add Python to PATH。
    pause
    exit /b 1
)

echo [1/4] 更新 pip...
python -m pip install --upgrade pip
if errorlevel 1 goto :failed

echo [2/4] 安装程序依赖...
python -m pip install -r requirements.txt
if errorlevel 1 goto :failed

echo [3/4] 安装 PyInstaller...
python -m pip install pyinstaller
if errorlevel 1 goto :failed

echo [4/4] 开始打包...
python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "语文试卷一键排版" ^
  word_exam_formatter.py

if errorlevel 1 goto :failed

echo.
echo ==========================================
echo 打包成功！
echo EXE 文件位置：
echo %~dp0dist\语文试卷一键排版.exe
echo ==========================================
explorer "%~dp0dist"
pause
exit /b 0

:failed
echo.
echo [失败] 打包没有完成，请查看上方错误信息。
pause
exit /b 1

