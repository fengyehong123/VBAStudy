' ====================================
' 🔷基础函数的用法
' Function 函数名(参数 As 类型) As 返回类型
'   函数名 = 返回值
' End Function
'
' vba中的函数是需要有返回值的, 如果没有返回值
' 那就不要使用函数, 而是使用 Sub
' ====================================

' 🔷定义一个基本的函数
Function Add(a As Double, b As Double) As Double
    Add = a + b
End Function

' 🔷可选参数
Function GetMsg(str1 As String, Optional str2 As String = "世界") As String
    GetMsg = "你好, " & str1 & ", " & str2
End Function

' 🔷判断是否为偶数
Function IsEven(n As Long) As Boolean
    IsEven = (n Mod 2 = 0)
End Function

' 🔷提前return函数
Function GetLevel(score As Integer) As String

    If score < 60 Then
        GetLevel = "不及格"
        ' 🔷提前return函数
        Exit Function
    End If
    
    If score < 80 Then
        GetLevel = "及格"
        Exit Function
    End If
    
    GetLevel = "优秀"

End Function

Sub MainSub()
    ' 调用函数, 获取返回值
    Debug.Print Add(1, 2)
    Debug.Print GetMsg("哈哈")
End Sub