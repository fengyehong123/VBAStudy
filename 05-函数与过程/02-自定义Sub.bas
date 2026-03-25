Sub MainSub()

    ' 调用Sub过程
    Call SayHello("张三")

    ' Call可以省略
    Greet("李四")

    ' 还可以这么使用
    Greet "王五"

    ' 定义一个数字
    Dim num1 As Long: num1 = 5

    ' 值传递调用
    Call Test1(num1)
    Debug.Print num1  ' 5

    ' 引用传递调用, 可以通过这种方式变相的获取返回值
    Call Test2(num1)
    Debug.Print num1  ' 15

End Sub

' 🔷定义一个Sub过程
private Sub SayHello(name As String)
    MsgBox "Hello, " & name
End Sub

' 🔷带有默认值的参数
private Sub Greet(name As String, Optional msg As String = "Hello")
    MsgBox msg & ", " & name
End Sub

' 🔷值传递参数
private Sub Test1(ByVal x As Long)
    x = x + 10
End Sub

' 🔷引用传递参数
private Sub Test2(ByRef x As Long)
    x = x + 10
End Sub

' 🔷提前结束Sub
private Sub HitMsg(score As Integer)

    If score < 60 Then
        MsgBox "不及格"
        Exit Sub
    End If
    
    If score < 80 Then
        MsgBox "及格"
        Exit Sub
    End If
    
    MsgBox "优秀"

End Sub