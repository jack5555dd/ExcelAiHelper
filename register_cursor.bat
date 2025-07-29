@echo off
echo Registering Cursor Excel Add-in...

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
set "DLL_PATH=%CURRENT_DIR%bin\Debug\CursorExcelAddin.dll"

rem 检查DLL文件是否存在
if not exist "%DLL_PATH%" (
    echo ERROR: Could not find %DLL_PATH%
    echo Please build the project first.
    pause
    exit /b 1
)

echo DLL Path: %DLL_PATH%

rem 尝试卸载现有注册
echo Unregistering any existing installations...
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /unregister
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /unregister

rem 注册DLL (32位)
echo Registering COM component for 32-bit Excel...
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase /tlb

rem 注册DLL (64位)
echo Registering COM component for 64-bit Excel...
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase /tlb

rem 导入注册表项
echo Adding registry entries...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\CursorExcelAddin.Connect" /v Description /t REG_SZ /d "Cursor Excel Add-in for Microsoft Excel" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\CursorExcelAddin.Connect" /v FriendlyName /t REG_SZ /d "Cursor Excel Add-in" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\CursorExcelAddin.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f

if %errorlevel% neq 0 (
    echo ERROR: Failed to add registry entries!
    pause
    exit /b 1
) else (
    echo Registry entries added successfully.
)

echo.
echo Cursor Excel Add-in registration complete!
echo Please restart Microsoft Excel to load the add-in.
pause 