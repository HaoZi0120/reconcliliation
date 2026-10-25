@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/3] 安装依赖...
python -m pip install -r requirements.txt
if errorlevel 1 goto :fail

echo [2/3] 清理旧构建...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [3/3] 打包（目录模式，便于公司电脑运行）...
python -m PyInstaller --noconfirm --clean "订单发票比对工具.spec"
if errorlevel 1 goto :fail

copy /Y "使用说明.txt" "dist\订单发票比对工具\使用说明.txt" >nul

echo.
echo 打包完成。请将整个文件夹拷贝到目标电脑：
echo   dist\订单发票比对工具\
echo 双击其中的「订单发票比对工具.exe」即可，无需安装 Python。
echo.
pause
exit /b 0

:fail
echo 打包失败。
pause
exit /b 1
