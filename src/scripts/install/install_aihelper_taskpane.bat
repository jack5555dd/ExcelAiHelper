@echo off
echo ==================================================
echo AIHelper Excel加载项安装程序 - 任务面板版
echo ==================================================

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

echo [步骤 1/6] 关闭Excel（如果正在运行）...
taskkill /f /im excel.exe >nul 2>&1
echo Excel已关闭（或未运行）。

echo [步骤 2/6] 卸载现有COM组件...
if exist "%DLL_PATH%" (
    "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /unregister /silent
    "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /unregister /silent
    echo 现有COM组件已卸载。
)

echo [步骤 3/6] 移除注册表项...
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\CursorExcelAddin.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\MsExcelAddin.Connect" /f >nul 2>&1
echo 注册表项已移除。

echo [步骤 4/6] 清理构建产物...
if exist "bin" rmdir /S /Q bin
if exist "obj" rmdir /S /Q obj
echo 构建产物已清理。

echo [步骤 5/6] 重新构建项目...
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" MsExcelAddin.csproj /p:Configuration=Debug /p:Platform="AnyCPU" /t:Rebuild

if %errorlevel% neq 0 (
    echo 构建失败！请检查编译错误。
    pause
    exit /b 1
)
echo 项目重新构建成功。

echo [步骤 6/6] 注册COM组件和加载项...
set "DLL_PATH=%CURRENT_DIR%bin\Debug\AIHelper.dll"
set "CONFIG_PATH=%CURRENT_DIR%bin\Debug\AIHelper.config"

rem 创建默认配置
echo ApiKey=> "%CONFIG_PATH%"
echo ApiEndpoint=https://api.openai.com/v1/chat/completions>> "%CONFIG_PATH%"
echo ApiModel=gpt-3.5-turbo>> "%CONFIG_PATH%"

rem 使用RegAsm注册
echo 使用RegAsm注册组件...
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase

rem 添加注册表项
echo 添加注册表项...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v Description /t REG_SZ /d "AI Helper Add-in for Microsoft Excel" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v FriendlyName /t REG_SZ /d "AIHelper Add-in" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f
echo COM组件已注册并添加注册表项。

echo.
echo ==================================================
echo 安装完成！
echo ==================================================
echo.
echo 重要说明:
echo 1. 启动Microsoft Excel
echo 2. 点击菜单栏中的"AIHelper"菜单
echo 3. 点击"显示AI助手"以显示右侧任务面板
echo 4. 点击"API设置"配置您的API密钥
echo.
echo 如果加载项未正确加载:
echo - 检查Excel的信任中心设置
echo - 确保已启用所有COM加载项
echo - 如有需要，再次运行此脚本
echo.
pause 