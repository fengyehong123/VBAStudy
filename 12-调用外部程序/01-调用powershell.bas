' ==========================================================
' 💥通过宏脚本执行powershell脚本有可能被杀毒软件提示
' 被提示的话, 需要选择允许
' ==========================================================
Sub RunPSAndGetOutput()

    Dim objShell As Object
    Dim objExec As Object
    
    ' powershell命令
    const psCmd = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command ""Get-Date"""

    ' WScript对象
    Set objShell = CreateObject("WScript.Shell")

    ' 🔷方式1: 调用
    ' 执行powershell的时候会有弹窗出现
    Set objExec = objShell.Exec(psCmd)
    ' 读取输出流
    Dim result As String: result = objExec.StdOut.ReadAll
    MsgBox result

    ' 🔷方式2: 调用
    ' 如果不需要获取返回值的话, 也可以直接使用 .Run
    ' 0 隐藏窗口, 后台执行
    ' 1 显示窗口
    objShell.Run psCmd, 0, True

    ' 🔷方式3: 调用
    ' vbHide 隐藏窗口, 后台执行
    ' vbNormalFocus 显示窗口
    Shell psCmd, vbHide
    
End Sub