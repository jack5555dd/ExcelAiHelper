Option Explicit

' 测试AIHelper加载项的Ribbon回调
Sub TestRibbonCallbacks()
    On Error Resume Next
    
    Dim objExcel, objAddin
    Dim strGUID, strProgID
    
    ' 设置GUID和ProgID
    strGUID = "{A2F47820-C9D3-4C8F-B3D5-78A982F89E31}"
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
        WScript.Echo " - " & addin.Description & " (" & addin.ProgId & ") 状态: " & addin.Connect
    Next
    
    ' 尝试找到我们的加载项
    Dim bFound
    bFound = False
    For Each addin In objExcel.COMAddIns
        If addin.ProgId = strProgID Then
            bFound = True
            Set objAddin = addin
            WScript.Echo "找到目标加载项: " & addin.Description & " 状态: " & addin.Connect
            Exit For
        End If
    Next
    
    If Not bFound Then
        WScript.Echo "未找到加载项 " & strProgID & "。检查是否已正确注册。"
    Else
        ' 确保加载项已连接
        If Not objAddin.Connect Then
            WScript.Echo "加载项未连接，尝试连接..."
            objAddin.Connect = True
            WScript.Sleep 1000
        End If
        
        ' 测试各种回调
        WScript.Echo "正在测试Ribbon回调..."
        
        ' 获取加载项实例
        Dim obj
        On Error Resume Next
        Set obj = GetObject(, strProgID)
        
        If Err.Number <> 0 Then
            WScript.Echo "无法获取加载项实例: " & Err.Description
        Else
            WScript.Echo "成功获取加载项实例"
            
            ' 调用About方法
            On Error Resume Next
            WScript.Echo "调用OnAbout方法..."
            ' 由于无法直接创建IRibbonControl对象，我们会看到错误
            ' 但这至少能验证方法是否存在
            obj.OnAbout
            
            If Err.Number <> 0 Then
                WScript.Echo "调用OnAbout出错: " & Err.Description
            Else
                WScript.Echo "OnAbout调用成功"
            End If
        End If
    End If
    
    WScript.Echo "测试完成。请在Excel中手动测试Ribbon按钮。"
    WScript.Echo "保持Excel打开以进行手动测试。"
    
    Set objAddin = Nothing
    ' 不关闭Excel，让用户可以手动测试
    ' objExcel.Quit
    Set objExcel = Nothing
End Sub

' 执行测试
TestRibbonCallbacks() 