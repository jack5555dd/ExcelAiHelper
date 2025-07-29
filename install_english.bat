@echo off
echo AIHelper Excel Add-in Installation
echo ===============================

rem Check for admin rights
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo ERROR: Please run this script as Administrator!
    echo Right-click on this file and select "Run as administrator".
    pause
    exit /b 1
)

rem Get current directory
set "CURRENT_DIR=%~dp0"
set "DLL_PATH=%CURRENT_DIR%bin\Debug\AIHelper.dll"
set "CONFIG_PATH=%CURRENT_DIR%bin\Debug\AIHelper.config"

echo [Step 1/4] Removing old versions...
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /f >nul 2>&1

rem Check if build is needed
if not exist "%DLL_PATH%" (
    echo [Step 2/4] Building project...
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" MsExcelAddin.csproj /p:Configuration=Debug /p:Platform="AnyCPU"
    
    if %errorlevel% neq 0 (
        echo Build failed! Please check compilation errors.
        pause
        exit /b 1
    )
)

echo [Step 3/4] Creating default configuration...
echo ApiKey=> "%CONFIG_PATH%"
echo ApiEndpoint=https://api.openai.com/v1/chat/completions>> "%CONFIG_PATH%"
echo ApiModel=gpt-3.5-turbo>> "%CONFIG_PATH%"

echo [Step 4/4] Registering COM component and add-in...
echo DLL Path: %DLL_PATH%

rem Register with RegAsm (.NET assembly)
echo Using RegAsm to register component...
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    echo Using 64-bit RegAsm to register component...
    "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase
)

echo Adding registry entries...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v Description /t REG_SZ /d "AI Helper Add-in for Microsoft Excel" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v FriendlyName /t REG_SZ /d "AIHelper Add-in" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f

echo.
echo Installation complete!
echo.
echo Important notes: 
echo 1. Please restart Microsoft Excel to load the add-in
echo 2. In Excel, click on the "AIHelper" tab
echo 3. Click "API Settings" button to configure your API key
echo 4. Click "Show AI Assistant" to display the AI assistant panel
echo.
pause 