@echo off
echo AIHelper Add-in Setup

rem 检查管理员权限
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo ERROR: Please run this script as Administrator!
    echo Right-click on this file and select "Run as administrator".
    pause
    exit /b 1
)

echo [Step 1/4] Cleaning up old installations...
rem 清除旧的注册表项
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\CursorExcelAddin.Connect" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /f >nul 2>&1

rem 清除旧的编译文件
if exist "bin" rmdir /S /Q bin
if exist "obj" rmdir /S /Q obj

echo [Step 2/4] Building the project...
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" MsExcelAddin.csproj /p:Configuration=Debug /p:Platform="AnyCPU" /t:Rebuild

if %errorlevel% neq 0 (
  echo Build failed!
  pause
  exit /b 1
)

echo [Step 3/4] Registering the COM component...
set "CURRENT_DIR=%~dp0"
set "DLL_PATH=%CURRENT_DIR%bin\Debug\AIHelper.dll"

"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase /tlb
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "%DLL_PATH%" /codebase /tlb

echo [Step 4/4] Setting up registry entries...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v Description /t REG_SZ /d "AI Helper Add-in for Microsoft Excel" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v FriendlyName /t REG_SZ /d "AIHelper Add-in" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f

echo.
echo AIHelper Add-in installation complete!
echo.
echo Important: Please restart Microsoft Excel to load the add-in.
echo The add-in should appear as a new tab called "AIHelper" in Excel's ribbon.
echo.
echo If you encounter any issues:
echo 1. Make sure Excel is completely closed before running this script
echo 2. Check Excel's Trust Center settings to ensure COM add-ins are allowed
echo 3. Try running this setup script again

pause 