using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;

// WPS ET COM 加载项主类
[ComVisible(true)]
[Guid("12345678-1234-1234-1234-123456789ABC")]
[ProgId("TestCompany.WpsEtAddin")]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class Connect
{
    private dynamic _app; // WPS ET Application 对象
    private dynamic _addInInst;

    /// <summary>
    /// 连接到 WPS ET 时调用
    /// </summary>
    public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
    {
        try
        {
            _app = application; // ET Application
            _addInInst = addInInst;
            
            // 初始化逻辑
            System.Windows.Forms.MessageBox.Show("WPS ET 加载项已成功加载！", "测试加载项");
            
            // 可以在这里添加菜单、按钮等
            // AddCustomMenu();
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show("加载项连接失败: " + ex.Message, "错误");
        }
    }

    /// <summary>
    /// 断开连接时调用
    /// </summary>
    public void OnDisconnection(int disconnectMode, ref Array custom)
    {
        try
        {
            // 清理资源
            _app = null;
            _addInInst = null;
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show("断开连接时出错: " + ex.Message, "错误");
        }
    }

    public void OnAddInsUpdate(ref Array custom)
    {
        // 加载项更新时调用
    }

    public void OnStartupComplete(ref Array custom)
    {
        // WPS 启动完成时调用
    }

    public void OnBeginShutdown(ref Array custom)
    {
        // WPS 开始关闭时调用
    }

    /// <summary>
    /// 添加自定义菜单示例（可选）
    /// </summary>
    private void AddCustomMenu()
    {
        try
        {
            // 这里可以添加自定义菜单的代码
            // 具体实现需要根据 WPS API 文档
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show("添加菜单失败: " + ex.Message, "错误");
        }
    }

    /// <summary>
    /// COM 注册方法
    /// </summary>
    [ComRegisterFunction]
    public static void RegisterFunction(Type type)
    {
        try
        {
            // 可以在这里添加额外的注册逻辑
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show("注册时出错: " + ex.Message, "注册错误");
        }
    }

    /// <summary>
    /// COM 注销方法
    /// </summary>
    [ComUnregisterFunction]
    public static void UnregisterFunction(Type type)
    {
        try
        {
            // 可以在这里添加额外的注销逻辑
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show("注销时出错: " + ex.Message, "注销错误");
        }
    }
}