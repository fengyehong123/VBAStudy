Sub MainSub()

    Dim a As Long: a = 10

    ' =============
    ' 🔷判断运算符
    ' =============
    ' 等于
    Debug.Print a = 10  ' True
    ' 不等于
    Debug.Print a <> 5  ' True

    ' 大于 
    Debug.Print a > 3  ' True
    ' 大于等于
    Debug.Print a >= 3  ' True
    ' 小于
    Debug.Print a < 3  ' False
    ' 小于等于
    Debug.Print a <= 3  ' False

    ' =============
    ' 🔷逻辑运算符
    ' =============
    Debug.Print a > 5 And a < 20  ' True
    Debug.Print a < 5 Or a = 10  ' True
    Debug.Print Not(a = 10)  ' False
    
End Sub