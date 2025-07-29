@echo off
echo AIHelper Add-in Manual Installation

rem 检查管理员权限
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo ERROR: Please run this script as Administrator!
    echo Right-click on this file and select "Run as administrator".
    pause
    exit /b 1
)

rem 获取当前目录和DLL路径
set "CURRENT_DIR=%~dp0"
set "DLL_PATH=%CURRENT_DIR%bin\Debug\AIHelper.dll"

echo [Step 1/3] Cleaning registry entries...
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\CursorExcelAddin.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /f >nul 2>&1

rem 检查DLL文件是否存在
if not exist "%DLL_PATH%" (
    echo ERROR: Could not find %DLL_PATH%
    echo Please build the project first with: msbuild MsExcelAddin.csproj
    pause
    exit /b 1
)

echo [Step 2/3] Registering COM component...
echo DLL Path: %DLL_PATH%

rem 使用RegAsm注册 (.NET 程序集不能用regsvr32注册)
echo Using RegAsm to register assembly...
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase /tlb
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    echo Using 64-bit RegAsm to register assembly...
    "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase /tlb
)

echo [Step 3/3] Adding registry entries...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v Description /t REG_SZ /d "AI Helper Add-in for Microsoft Excel" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v FriendlyName /t REG_SZ /d "AIHelper Add-in" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f

echo.
echo AIHelper Add-in installation complete!
echo.
echo Important: Please restart Microsoft Excel to load the add-in.
echo The add-in should appear as a new tab called "AIHelper" in Excel's ribbon.
echo.
pause 