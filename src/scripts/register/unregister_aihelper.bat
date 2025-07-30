@echo off
echo Unregistering AIHelper Add-in...

rem 检查管理员权限
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo ERROR: Please run this script as Administrator!
    echo Right-click on this file and select "Run as administrator".
    pause
    exit /b 1
)

rem 获取当前目录
set "CURRENT_DIR=%~dp0"
set "DLL_PATH=%CURRENT_DIR%bin\Debug\AIHelper.dll"

rem 检查DLL文件是否存在
if exist "%DLL_PATH%" (
    rem 卸载注册DLL
    echo Unregistering COM component...
    "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /unregister
    "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /unregister
    if %errorlevel% neq 0 (
        echo WARNING: COM component unregistration failed! Continuing to delete registry entries.
    ) else (
        echo COM component unregistered successfully.
    )
) else (
    echo WARNING: Could not find %DLL_PATH%
    echo Continuing to delete registry entries.
)

rem 删除注册表项
echo Deleting registry entries...
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /f
if %errorlevel% neq 0 (
    echo WARNING: Registry entries deletion failed! They might not exist or were already deleted.
) else (
    echo Registry entries deleted successfully.
)

echo.
echo AIHelper Add-in unregistration complete!
echo Please restart Microsoft Excel to apply changes.
pause 