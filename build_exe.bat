@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/4] 安装依赖...
python -m pip install -r requirements.txt
if errorlevel 1 goto :fail

REM 从 reconcile_gui.py 读取 APP_VERSION，输出目录自动带版本号
for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command "(Select-String -Path reconcile_gui.py -Pattern 'APP_VERSION\s*=\s*\"([^\"]+)\"').Matches.Groups[1].Value"`) do set APP_VER=%%i
if "%APP_VER%"=="" set APP_VER=unknown
set PACK_NAME=订单发票比对工具_%APP_VER%
echo 版本：%APP_VER%
echo 输出目录：dist\%PACK_NAME%\

echo [2/4] 清理本次构建缓存...
if exist build rmdir /s /q build
if exist "dist\_build_staging" rmdir /s /q "dist\_build_staging"

REM 旧无版本目录若被占用则挪走，避免 PyInstaller 清目录失败
if exist "dist\订单发票比对工具" (
  powershell -NoProfile -Command "try { Move-Item -LiteralPath 'dist\订单发票比对工具' -Destination ('dist\_locked_old_' + (Get-Date -Format 'yyyyMMdd_HHmmss')) -Force } catch { }"
)

echo [3/4] 打包（目录模式，便于公司电脑运行）...
python -m PyInstaller --noconfirm --clean --distpath "dist\_build_staging" "订单发票比对工具.spec"
if errorlevel 1 goto :fail

copy /Y "使用说明.txt" "dist\_build_staging\订单发票比对工具\使用说明.txt" >nul

echo [4/4] 生成带版本号目录与压缩包...
if exist "dist\%PACK_NAME%" rmdir /s /q "dist\%PACK_NAME%"
xcopy /E /I /Y "dist\_build_staging\订单发票比对工具" "dist\%PACK_NAME%" >nul
if exist "dist\%PACK_NAME%.zip" del /F /Q "dist\%PACK_NAME%.zip"
powershell -NoProfile -Command "Compress-Archive -Path 'dist\%PACK_NAME%' -DestinationPath 'dist\%PACK_NAME%.zip' -Force"
if errorlevel 1 goto :fail

rmdir /s /q "dist\_build_staging" 2>nul

echo.
echo 打包完成。请将整个文件夹拷贝到目标电脑：
echo   dist\%PACK_NAME%\
echo 或解压：
echo   dist\%PACK_NAME%.zip
echo 双击其中的「订单发票比对工具.exe」即可，无需安装 Python。
echo.
pause
exit /b 0

:fail
echo 打包失败。
pause
exit /b 1
