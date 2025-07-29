@echo off
echo ===============================================================
echo AI Helper - Excel Add-in Complete Reinstallation
echo ===============================================================
echo.

echo [Step 1/7] Closing Excel (if running)...
taskkill /f /im excel.exe >nul 2>&1
echo Excel has been closed (or was not running).

echo [Step 2/7] Cleaning registry entries...
regedit /s cleanup_addins.reg
echo Registry entries cleaned.

echo [Step 3/7] Unregistering existing COM components...
if exist "bin\Debug\AIHelper.dll" (
    "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "bin\Debug\AIHelper.dll" /unregister /silent
    "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "bin\Debug\AIHelper.dll" /unregister /silent
    echo Existing COM components unregistered.
)

echo [Step 4/7] Cleaning build artifacts...
if exist "bin" rmdir /S /Q bin
if exist "obj" rmdir /S /Q obj
echo Build artifacts cleaned.

echo [Step 5/7] Rebuilding the project...
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" MsExcelAddin.csproj /p:Configuration=Debug /p:Platform=AnyCPU /t:Rebuild

if %errorlevel% neq 0 (
    echo Build failed! Please check compilation errors.
    pause
    exit /b 1
)
echo Project rebuilt successfully.

echo [Step 6/7] Registering COM component...
set "DLL_PATH=%~dp0bin\Debug\AIHelper.dll"
set "CONFIG_PATH=%~dp0bin\Debug\AIHelper.config"

rem Create default config
echo ApiKey=> "%CONFIG_PATH%"
echo ApiEndpoint=https://api.openai.com/v1/chat/completions>> "%CONFIG_PATH%"
echo ApiModel=gpt-3.5-turbo>> "%CONFIG_PATH%"

rem Register with RegAsm
echo Registering COM component...
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase

echo [Step 7/7] Adding registry entries...
echo Creating registry file...
echo Windows Registry Editor Version 5.00 > register_aihelper.reg
echo. >> register_aihelper.reg
echo [HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect] >> register_aihelper.reg
echo "Description"="AI Helper Add-in for Microsoft Excel" >> register_aihelper.reg
echo "FriendlyName"="AI Helper" >> register_aihelper.reg
echo "LoadBehavior"=dword:00000003 >> register_aihelper.reg

regedit /s register_aihelper.reg
echo Registry entries added.

echo.
echo ===============================================================
echo Installation complete!
echo ===============================================================
echo.
echo Important notes:
echo 1. Start Microsoft Excel
echo 2. Click on the "AI Helper" menu in the menu bar
echo 3. Click on "Show AI Assistant" to display the assistant
echo 4. Configure API settings by clicking "API Settings"
echo.
echo If the add-in fails to load:
echo - Check Excel Trust Center settings
echo - Make sure all COM add-ins are enabled
echo - Run this script again if needed
echo.
pause 