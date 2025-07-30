On Error Resume Next

Dim xl
Set xl = CreateObject("Excel.Application")
xl.Visible = True
WScript.Echo "Excel已启动，请检查加载项。"

' 保持Excel打开
Set xl = Nothing 