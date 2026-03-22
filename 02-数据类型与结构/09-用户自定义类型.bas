' ================================
' 🔷定义一个自定义类型
' 在 VBA 中：
'   1. Type 只能定义在模块级(Module Level)
'   2. 不能写在过程(Sub / Function)内部
'   3. 可以写在Sub的外部, 但是需要使用 private 限制作用域
' ================================
private Type Person
    Name As String
    Age As Long
    Address As String
End Type

Sub MainSub()

    ' 🔷创建自定义类型并赋值
    Dim p1 As Person
    p1.Name = "Tom"
    p1.Age = 20
    p1.Address = "地球"

    ' 🔷创建自定义类型并赋值
    '    使用了With块, 简化书写
    Dim p2 As Person
    With p2
        .Name = "贾飞天"
        .Age = 23
        .Address = "月球"
    End With

    Debug.Print p1.Name  ' Tom
    Debug.Print p2.Name  ' 贾飞天
End Sub