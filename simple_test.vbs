Option Explicit

' 简单测试AIHelper加载项
Sub TestAddin()
    On Error Resume Next
    
    Dim objExcel
    Dim strProgID
    
    strProgID = "AIHelper.Connect"
    
    ' 创建Excel实例
    WScript.Echo "正在启动Excel..."
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    
    ' 等待Excel启动
    WScript.Sleep 2000
    
    ' 显示Excel版本
    WScript.Echo "Excel版本: " & objExcel.Version
    
    ' 检查加载项列表
    Dim addin
    WScript.Echo "已加载的加载项列表:"
    For Each addin In objExcel.COMAddIns
        WScript.Echo " - " & addin.Description & " (" & addin.ProgId & "), 状态: " & addin.Connect
    Next
    
    WScript.Echo "请在Excel中手动测试Ribbon按钮。测试完成后关闭Excel。"
    Dim tempPath
    tempPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%")
    WScript.Echo "加载项调试信息保存在: " & tempPath & "\VSTOAddinInfo.log"
    
    ' 不关闭Excel，让用户可以手动测试
    Set objExcel = Nothing
End Sub

' 执行测试
TestAddin() 