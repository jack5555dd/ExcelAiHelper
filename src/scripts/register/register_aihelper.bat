@echo off
echo Registering AIHelper Add-in...

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
if not exist "%DLL_PATH%" (
    echo ERROR: Could not find %DLL_PATH%
    echo Please build the project first.
    pause
    exit /b 1
)

echo DLL Path: %DLL_PATH%

rem 先尝试卸载任何已存在的注册
echo Unregistering any existing installations...
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\CursorExcelAddin.Connect" /f >nul 2>&1
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /unregister >nul 2>&1
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /unregister >nul 2>&1

rem 注册DLL (32位和64位)
echo Registering COM component...
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase /tlb
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase /tlb

rem 导入注册表项
echo Adding registry entries...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v Description /t REG_SZ /d "AI Helper Add-in for Microsoft Excel" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v FriendlyName /t REG_SZ /d "AIHelper Add-in" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f

if %errorlevel% neq 0 (
    echo ERROR: Failed to add registry entries!
    pause
    exit /b 1
) else (
    echo Registry entries added successfully.
)

echo.
echo AIHelper Add-in registration complete!
echo Please restart Microsoft Excel to load the add-in.
pause 