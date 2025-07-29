@echo off
echo 检查AIHelper加载项注册情况
echo ==========================

echo 检查Excel加载项注册表...
reg query "HKCU\Software\Microsoft\Office\Excel\Addins\AIHelper.Connect"
if %errorlevel% neq 0 (
    echo [错误] 注册表中未找到AIHelper.Connect加载项!
) else (
    echo [成功] 注册表中找到AIHelper.Connect加载项。
)

echo.
echo 检查CLSID注册...
reg query "HKCR\CLSID\{A2F47820-C9D3-4C8F-B3D5-78A982F89E31}" /s
if %errorlevel% neq 0 (
    echo [错误] 注册表中未找到CLSID {A2F47820-C9D3-4C8F-B3D5-78A982F89E31}!
) else (
    echo [成功] 注册表中找到CLSID。
)

echo.
echo 检查ProgID注册...
reg query "HKCR\AIHelper.Connect" /s
if %errorlevel% neq 0 (
    echo [错误] 注册表中未找到ProgID AIHelper.Connect!
) else (
    echo [成功] 注册表中找到ProgID。
)

echo.
echo 检查完成。

pause 