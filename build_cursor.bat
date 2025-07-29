@echo off
echo Building Cursor Excel Add-in...
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" MsExcelAddin.csproj /p:Configuration=Debug /p:Platform="AnyCPU" /t:Rebuild
if %errorlevel% neq 0 (
  echo Build failed!
  pause
  exit /b 1
)
echo Build successful!
pause 