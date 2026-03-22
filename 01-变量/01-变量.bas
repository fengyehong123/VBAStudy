' 👍推荐写法👍
' 在模块顶部写
'   防止拼写错误
'   强制写Dim
Option Explicit

Sub MainSub()

    ' vba中使用Dim定义变量
    Dim str1 As String
    ' 定义变量的时候初始化赋值
    Dim str2 As String: str2 = "你好"

    ' 一次定义多个变量, 注意: 每个变量都需要使用 As
    ' 定义了3个Long类型的变量
    Dim a As Long, b As Long, c As Long

    ' 如果我们像下面这么写的话, 中间省略了 As 的话
    '   a → Variant
    '   b → Variant
    '   c → Integer
    Dim d, e, f As Long
    
    ' 先定义一个字符串, 然后赋值
    Dim str3 As String
    str3 = "Hello"

    ' 动态数组
    Dim arr3() As Integer
    ' 动态数组必须用 ReDim 才能使用
    ReDim arr3(1 To 5)

End Sub

' =========================
' Static 变量
'   不会被释放
'   类似【记忆变量】
' =========================
Function Counter()
    ' 每次调用函数的时候, 被Static修饰的变量值并不会重置
    ' 因此当 Counter函数被多次调用的时候, count变量的值会不断递增
    Static count As Integer
    count = count + 1
    ' 给函数设置返回值
    Counter = count
End Function