' =========================================
' 🔷ParamArray 是什么？
'   可以让 Sub / Function 接收【任意个参数】
'   类似于
'       JavaScript：...args
'       Python：*args
'
' 💥注意点
'   1. 必须是 Variant 数组
'   2. 参数名后面必须是 ()
'   3. 当ParamArray和普通参数一起使用的时候, 要放在最后
'   4. 只能有一个 ParamArray
' =========================================
Sub MainSub()
    Debug.Print JoinText("你", "好", "啊")  ' 你好啊
    Call Test("你", "1", "2", "3", "4", "5")
    ' 你1
    ' 你2
    ' 你3
    ' 你4
    ' 你5
End Sub

' 🔷定义一个拼接字符串的函数
Function JoinText(ParamArray arr() As Variant) As String
    ' 判断有参数了之后才进行处理
    If UBound(arr) >= 0 Then
        Dim i As Integer
        For i = 0 To UBound(arr)
            JoinText = JoinText & arr(i)
        Next
    End If
End Function

' 🔷当ParamArray和普通参数一起使用的时候, 要放在最后
private Sub Test(prefix As String, ParamArray args() As Variant)
    Dim i As Integer
    For i = 0 To UBound(args)
        Debug.Print prefix & args(i)
    Next
End Sub