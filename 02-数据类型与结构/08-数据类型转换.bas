' =========================================
' 🔷数据类型转换的必要性
'   VBA 默认很多数据是 Variant, 例如
'       Dim v
'       v = Range("A1").Value
'   此时获取到的v的数据类型, 有可能是
'       数字
'       字符串
'       日期
'       空值
' =========================================
Sub MainSub()

    ' 🔷CInt 转 Integer
    Dim x1 As Integer
    ' 结果是 4 
    ' ※四舍五入
    x1 = CInt(3.6)  ' 4

    ' 🔷CLng 转 Long
    Dim x2 As Long
    x2 = CLng(3.6)  ' 4
    Dim x3 As Long: x3 = CLng("3.6")
    Debug.Print x3

    ' 🔷CDbl 转 浮点数
    Dim x4 As Double
    x4 = CDbl("123.45")

    ' 🔷CStr 转 字符串
    Dim x5 As String
    x5 = CStr(123)

    ' 🔷CBool 转 布尔值
    Dim b As Boolean
    b = CBool(1)  ' True
    b = CBool(0)  ' False
    
    ' 🔷CDate 转 日期
    Dim d As Date
    d = CDate("2025-03-21")

    ' 🔷CCur 转 Currency
    Dim m As Currency
    m = CCur(123.456)

    ' 🔷Val 函数
    '   从左往右读数字
    '   遇到非数字停止
    Dim x6 As Double: x6 = Val("123abc")
    Debug.Print x6  ' 123
    Debug.Print Val("12.3abc")  ' 12.3
    Debug.Print Val("abc123")  ' 0

    ' 定义一个可以存放任意数据类型的变量
    Dim v As Variant
    v = "123"

    ' ✅安全的转换写法
    ' 判断是否是数字
    If IsNumeric(v) Then
        Debug.Print CLng(v)
    End If

    ' 判断是否是日期
    If IsDate("2025-03-21") Then
        CDate("2025-03-21")
    End If

End Sub