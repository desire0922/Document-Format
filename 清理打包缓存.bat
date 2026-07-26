@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "语文试卷一键排版.spec" del /q "语文试卷一键排版.spec"
echo 已清理打包缓存。
pause
