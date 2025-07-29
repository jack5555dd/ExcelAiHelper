@echo off
echo 正在编译MsExcelAddin...
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" MsExcelAddin.csproj /p:Configuration=Debug /p:Platform="AnyCPU" /t:Rebuild
if %errorlevel% neq 0 (
  echo 编译失败！
  pause
  exit /b 1
)
echo 编译成功！
pause 