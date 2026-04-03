Sub MainSub()

    ' 显示状态栏的提示信息
    Application.StatusBar = "正在处理中..."

    ' 延时(替代 sleep)
    Application.Wait Now + TimeValue("00:00:05")

    ' 关闭状态栏的提示信息
    Application.StatusBar = False
    
End Sub