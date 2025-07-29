@echo off
echo 正在构建 Microsoft Excel 辅助功能加载项...

rem 检查是否安装了MSBuild
where msbuild >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到MSBuild工具。
    echo 请确保安装了Visual Studio或.NET Framework SDK。
    pause
    exit /B 1
)

rem 获取当前目录
set "CURRENT_DIR=%~dp0"
set "PROJECT_FILE=%CURRENT_DIR%MsExcelAddin.csproj"

rem 检查项目文件是否存在
if not exist "%PROJECT_FILE%" (
    echo 错误: 找不到项目文件 %PROJECT_FILE%
    pause
    exit /B 1
)

rem 构建项目
echo 正在构建项目...
msbuild "%PROJECT_FILE%" /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild

if %errorlevel% neq 0 (
    echo 错误: 项目构建失败!
    pause
    exit /B 1
) else (
    echo 项目构建成功。
)

echo.
echo Microsoft Excel 辅助功能加载项构建完成!
echo 现在可以运行 register.bat 注册加载项。
pause 