@echo off
echo 正在注册 Microsoft Excel 辅助功能加载项...

rem 检查管理员权限
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo 错误: 请以管理员身份运行此脚本！
    echo 右键点击此文件，选择"以管理员身份运行"。
    pause
    exit /b 1
)

rem 获取当前目录
set "CURRENT_DIR=%~dp0"
set "DLL_PATH=%CURRENT_DIR%bin\Debug\MsExcelAddin.dll"

rem 检查DLL文件是否存在
if not exist "%DLL_PATH%" (
    echo 错误: 找不到 %DLL_PATH%
    echo 请先编译项目，然后再运行此脚本。
    pause
    exit /b 1
)

echo DLL路径: %DLL_PATH%

rem 注册DLL
echo 正在使用RegAsm注册COM组件...
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase /tlb
if %errorlevel% neq 0 (
    echo 错误: COM组件注册失败!
    pause
    exit /b 1
) else (
    echo COM组件注册成功。
)

rem 导入注册表项
echo 正在导入注册表项...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\MsExcelAddin.Connect" /v Description /t REG_SZ /d "Microsoft Excel 辅助功能加载项" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\MsExcelAddin.Connect" /v FriendlyName /t REG_SZ /d "辅助功能加载项" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\MsExcelAddin.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f

if %errorlevel% neq 0 (
    echo 错误: 注册表导入失败!
    pause
    exit /b 1
) else (
    echo 注册表项导入成功。
)

echo.
echo Microsoft Excel 辅助功能加载项注册完成!
echo 请重新启动 Microsoft Excel 以加载此加载项。
pause 