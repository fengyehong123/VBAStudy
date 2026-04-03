Sub MainSub()

    ' =======================================
    ' 🔷普通输入框
    '   返回的类型是字符串
    '   当用户点击取消的时候, 返回的是空字符串
    ' =======================================
    Dim inputContent As String: inputContent = InputBox("请输入你的名字：", "输入")
    Debug.Print inputContent

    ' =======================================
    ' 🔷可以限制输入类型的输入框
    ' 1 数字
    ' 2 字符串
    ' 4 布尔值
    ' 8 Range
    ' =======================================
    ' 获取用户输入的数字
    Dim num As Double
    num = Application.InputBox("请输入一个数字：", Type:=1)
    Debug.Print num

    ' 获取用户选中的区域
    Dim rng As Range
    Set rng = Application.InputBox("请在当前Sheet页面选择一个区域:", Type:=8)
    MsgBox rng.Address

End Sub