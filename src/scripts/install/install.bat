@echo off
chcp 65001 >nul
echo ========================================
echo WPS ET COM 加载项安装脚本
echo ========================================
echo.

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [√] 已获得管理员权限
) else (
    echo [×] 需要管理员权限运行此脚本
    echo 请右键点击此文件，选择"以管理员身份运行"
    pause
    exit /b 1
)

:: 设置变量
set "ADDIN_NAME=WpsEtAddin"
set "PROG_ID=TestCompany.WpsEtAddin"
set "DLL_PATH=%~dp0bin\Debug\WpsEtAddin.dll"
set "REGASM_PATH=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
set "WPS_ADDIN_DIR=%LOCALAPPDATA%\Kingsoft\Office\*\office6\AddinsWL\%ADDIN_NAME%"

echo [1] 检查文件是否存在...
if not exist "%DLL_PATH%" (
    echo [×] 找不到 DLL 文件: %DLL_PATH%
    echo 请先编译项目生成 DLL 文件
    pause
    exit /b 1
)
echo [√] DLL 文件存在

if not exist "%REGASM_PATH%" (
    echo [×] 找不到 RegAsm.exe: %REGASM_PATH%
    echo 请确保已安装 .NET Framework 4.0 或更高版本
    pause
    exit /b 1
)
echo [√] RegAsm.exe 存在

echo.
echo [2] 注册 COM 组件...
"%REGASM_PATH%" "%DLL_PATH%" /codebase /tlb
if %errorLevel% == 0 (
    echo [√] COM 组件注册成功
) else (
    echo [×] COM 组件注册失败
    pause
    exit /b 1
)

echo.
echo [3] 创建 WPS 加载项目录...
:: 查找 WPS 安装目录
for /d %%i in ("%LOCALAPPDATA%\Kingsoft\Office\*") do (
    if exist "%%i\office6" (
        set "WPS_ADDIN_DIR=%%i\office6\AddinsWL\%ADDIN_NAME%"
        goto :found_wps
    )
)

:: 如果在 LOCALAPPDATA 找不到，尝试 ProgramFiles
for /d %%i in ("%ProgramFiles%\Kingsoft\Office\*") do (
    if exist "%%i\office6" (
        set "WPS_ADDIN_DIR=%%i\office6\AddinsWL\%ADDIN_NAME%"
        goto :found_wps
    )
)

echo [×] 找不到 WPS 安装目录
echo 请手动将 publish.xml 复制到 WPS 的 AddinsWL 目录
pause
exit /b 1

:found_wps
echo [√] 找到 WPS 目录: %WPS_ADDIN_DIR%

if not exist "%WPS_ADDIN_DIR%" (
    mkdir "%WPS_ADDIN_DIR%" 2>nul
    if %errorLevel% == 0 (
        echo [√] 创建加载项目录成功
    ) else (
        echo [×] 创建加载项目录失败
        pause
        exit /b 1
    )
) else (
    echo [√] 加载项目录已存在
)

echo.
echo [4] 复制文件...
copy /Y "%~dp0publish.xml" "%WPS_ADDIN_DIR%\" >nul
if %errorLevel% == 0 (
    echo [√] publish.xml 复制成功
) else (
    echo [×] publish.xml 复制失败
    pause
    exit /b 1
)

:: 创建 bin 目录并复制 DLL
if not exist "%WPS_ADDIN_DIR%\bin" mkdir "%WPS_ADDIN_DIR%\bin"
copy /Y "%DLL_PATH%" "%WPS_ADDIN_DIR%\bin\" >nul
copy /Y "%~dp0bin\Debug\WpsEtAddin.tlb" "%WPS_ADDIN_DIR%\bin\" >nul 2>nul
echo [√] DLL 文件复制成功

echo.
echo [5] 添加注册表项...
reg add "HKCU\Software\Kingsoft\Office\WPS\AddinsWL" /v "%PROG_ID%" /t REG_SZ /d "" /f >nul
if %errorLevel% == 0 (
    echo [√] 注册表项添加成功
) else (
    echo [!] 注册表项添加失败（可能不影响加载）
)

echo.
echo ========================================
echo [√] 安装完成！
echo ========================================
echo.
echo 安装位置: %WPS_ADDIN_DIR%
echo.
echo 请重启 WPS 表格，然后按 Alt+F12 查看加载日志
echo 如果加载成功，会弹出"WPS ET 加载项已成功加载！"的消息框
echo.
pause