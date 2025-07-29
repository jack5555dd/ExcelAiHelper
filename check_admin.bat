@echo off
echo 正在检查管理员权限...
echo.

>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

if '%errorlevel%' NEQ '0' (
    echo 【权限状态】: 当前不具有管理员权限
    echo 【建议操作】: 请右键点击此批处理文件，选择"以管理员身份运行"
    echo.
    echo 警告: 未获得管理员权限可能导致COM加载项注册失败！
) else (
    echo 【权限状态】: 当前已具有管理员权限
    echo 【注册状态】: 可以继续进行COM加载项注册
    echo.
    echo 提示: 您现在可以安全地运行register.bat或install.bat进行加载项安装
)

echo.
pause 