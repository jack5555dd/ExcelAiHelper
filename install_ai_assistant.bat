@echo off
chcp 936 >nul
echo AI助手Excel加载项安装程序
echo ====================

rem 检查管理员权限
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo 错误: 请以管理员身份运行此脚本！
    echo 请右键点击此文件，选择"以管理员身份运行"。
    pause
    exit /b 1
)

rem 获取当前目录
set "CURRENT_DIR=%~dp0"
set "DLL_PATH=%CURRENT_DIR%bin\Debug\AIHelper.dll"
set "CONFIG_PATH=%CURRENT_DIR%bin\Debug\AIHelper.config"

echo [步骤 1/4] 清除旧版本...
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /f >nul 2>&1

rem 检查是否需要编译
if not exist "%DLL_PATH%" (
    echo [步骤 2/4] 编译项目...
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" MsExcelAddin.csproj /p:Configuration=Debug /p:Platform="AnyCPU"
    
    if %errorlevel% neq 0 (
        echo 构建失败！请检查编译错误。
        pause
        exit /b 1
    )
)

echo [步骤 3/4] 创建默认配置...
echo ApiKey=> "%CONFIG_PATH%"
echo ApiEndpoint=https://api.openai.com/v1/chat/completions>> "%CONFIG_PATH%"
echo ApiModel=gpt-3.5-turbo>> "%CONFIG_PATH%"

echo [步骤 4/4] 注册COM组件和加载项...
echo DLL路径: %DLL_PATH%

rem 使用RegAsm注册 (.NET 程序集)
echo 使用RegAsm注册组件...
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    echo 使用64位RegAsm注册组件...
    "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase
)

echo 添加注册表项...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v Description /t REG_SZ /d "AI Helper Add-in for Microsoft Excel" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v FriendlyName /t REG_SZ /d "AIHelper Add-in" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f

echo.
echo 安装完成！
echo.
echo 重要提示: 
echo 1. 请重新启动Microsoft Excel以加载加载项
echo 2. 在Excel中，点击"AIHelper"选项卡
echo 3. 点击"API Settings"按钮配置您的API密钥
echo 4. 点击"Show AI Assistant"显示AI助手面板
echo.
pause 