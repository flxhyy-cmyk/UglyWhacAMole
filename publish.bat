@echo off
chcp 65001 >nul
echo ========================================
echo   WindowInspector 发布工具
echo ========================================
echo.

echo [1/4] 清理旧的发布文件...
if exist "publish" rmdir /s /q "publish"
if exist "WindowInspector-Release.zip" del /f /q "WindowInspector-Release.zip"

echo [2/4] 编译并发布程序...
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o publish

if %errorlevel% neq 0 (
    echo.
    echo ❌ 发布失败！
    pause
    exit /b 1
)

echo [3/4] 复制必要文件...
copy README.md publish\ >nul 2>&1
copy build.bat publish\ >nul 2>&1
copy run.bat publish\ >nul 2>&1

echo [4/4] 创建发布包...
powershell -Command "Compress-Archive -Path 'publish\*' -DestinationPath 'WindowInspector-Release.zip' -Force"

echo.
echo ========================================
echo   ✅ 发布完成！
echo ========================================
echo.
echo 发布文件位置:
echo   - 程序目录: publish\
echo   - 压缩包: WindowInspector-Release.zip
echo.
echo 可执行文件: publish\WindowInspector.exe
echo.
echo 环境要求:
echo   - .NET 8.0 Runtime
echo.
echo 提示: 
echo   - 如需完全独立版本，请修改 publish.bat 中的 --self-contained 为 true
echo.
pause
