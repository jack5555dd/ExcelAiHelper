@echo off
echo 正在卸载 Microsoft Excel 辅助功能加载项...

rem 获取管理员权限
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo 请求管理员权限...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"

rem 获取当前目录
set "CURRENT_DIR=%~dp0"
set "DLL_PATH=%CURRENT_DIR%bin\Debug\MsExcelAddin.dll"

rem 检查DLL文件是否存在
if not exist "%DLL_PATH%" (
    echo 警告: 找不到 %DLL_PATH%
    echo 将尝试继续卸载注册表项。
) else (
    rem 卸载注册DLL
    echo 正在卸载COM组件...
    regsvr32 /u /s "%DLL_PATH%"
    if %errorlevel% neq 0 (
        echo 警告: COM组件卸载失败! 尝试继续删除注册表项。
    ) else (
        echo COM组件卸载成功。
    )
)

rem 删除注册表项
echo 正在删除注册表项...
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\MsExcelAddin.Connect" /f
if %errorlevel% neq 0 (
    echo 警告: 注册表项删除失败！可能不存在或已被删除。
) else (
    echo 注册表项删除成功。
)

echo.
echo Microsoft Excel 辅助功能加载项卸载完成!
echo 请重新启动 Microsoft Excel 以应用更改。
pause 