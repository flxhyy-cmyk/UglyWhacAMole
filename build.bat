@echo off
chcp 65001 >nul
echo ========================================
echo   WindowInspector 编译工具
echo ========================================
echo.

echo [1/3] 清理旧的编译文件...
dotnet clean -c Release >nul 2>&1

echo [2/3] 正在编译项目...
dotnet build -c Release

if %errorlevel% == 0 (
    echo.
    echo [3/3] 编译完成
    echo.
    echo ========================================
    echo   ✓ 编译成功！
    echo ========================================
    echo.
    echo 可执行文件: bin\Release\net8.0-windows\WindowInspector.exe
    echo.
) else (
    echo.
    echo ========================================
    echo   ✗ 编译失败
    echo ========================================
    echo.
    echo 请检查上方的错误信息
    echo.
)

pause
