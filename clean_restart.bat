@echo off
echo ===================================================
echo Complete Cleanup and Reinstallation of AIHelper Add-in
echo ===================================================

rem Check for admin rights
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo ERROR: Please run this script as Administrator!
    echo Right-click on this file and select "Run as administrator".
    pause
    exit /b 1
)

echo [Step 1/6] Closing Excel (if running)...
taskkill /f /im excel.exe >nul 2>&1
echo Excel closed (or was not running).

echo [Step 2/6] Unregistering any existing COM components...
set "CURRENT_DIR=%~dp0"
set "DLL_PATH=%CURRENT_DIR%bin\Debug\AIHelper.dll"

if exist "%DLL_PATH%" (
    "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /unregister /silent
    "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /unregister /silent
    echo Existing COM component unregistered.
)

echo [Step 3/6] Removing registry entries...
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\CursorExcelAddin.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\MsExcelAddin.Connect" /f >nul 2>&1
echo Registry entries removed.

echo [Step 4/6] Cleaning build artifacts...
if exist "bin" rmdir /S /Q bin
if exist "obj" rmdir /S /Q obj
echo Build artifacts cleaned.

echo [Step 5/6] Rebuilding project...
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" MsExcelAddin.csproj /p:Configuration=Debug /p:Platform="AnyCPU" /t:Rebuild

if %errorlevel% neq 0 (
    echo Build failed! Please check compilation errors.
    pause
    exit /b 1
)
echo Project rebuilt successfully.

echo [Step 6/6] Registering COM component and add-in...
set "DLL_PATH=%CURRENT_DIR%bin\Debug\AIHelper.dll"
set "CONFIG_PATH=%CURRENT_DIR%bin\Debug\AIHelper.config"

rem Create default config
echo ApiKey=> "%CONFIG_PATH%"
echo ApiEndpoint=https://api.openai.com/v1/chat/completions>> "%CONFIG_PATH%"
echo ApiModel=gpt-3.5-turbo>> "%CONFIG_PATH%"

rem Register with RegAsm
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase
)

rem Add registry entries
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v Description /t REG_SZ /d "AI Helper Add-in for Microsoft Excel" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v FriendlyName /t REG_SZ /d "AIHelper Add-in" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f
echo COM component registered and registry entries added.

echo.
echo ===================================================
echo Installation complete!
echo ===================================================
echo.
echo Important steps:
echo 1. Start Microsoft Excel
echo 2. Look for the "AIHelper" tab in the ribbon
echo 3. Click "Show AI Assistant" to display the AI panel
echo 4. Configure your API key in "API Settings"
echo.
echo If the add-in does not load correctly:
echo - Check Excel's Trust Center settings
echo - Ensure all COM add-ins are enabled
echo - If needed, run this script again
echo.
pause 