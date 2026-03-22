' ===============================================
' 💥注意事项
'   插入类模块之后, 默认的名称叫【类1】
'   无法通过右键或者F2来改名
' 
' 🔷解决办法
'   点击默认生成的类名, 然后按下键盘上的F4键调出属性面板
'   然后找到名称, 再然后就可以改名了
' ===============================================
' 假设创建了一个类模块, 名称叫做Person
' 类的属性
Public Name As String
Public Age As Long

' 类中的方法
Public Sub SayHello()
    Debug.Print "Hello, I am " & Name
End Sub

' ===============================================
' 定义一个Sub调用类
Sub Test()

    ' =======================================================
    ' 和其他语言不同的是, 在使用类模块的时候并不需要进行导入
    ' 直接New使用即可
    ' =======================================================

    ' 创建类对象, 并给属性赋值
    Dim p1 As New Person
    p1.Name = "贾飞天"
    p1.Age = 23

    ' 调用类方法
    p1.SayHello

    ' 创建类对象, 通过With代码块赋值
    Dim p2 As New Person
    With p2
        .Name = "张三"
        .Age = 40
    End With
    p2.SayHello

    ' 定义一个用来存放类属性的数组, 然后初始化
    ' 注意: Array中的元素类型是Variant
    Dim props As Variant: props = Array("Name", "Age")

    Dim i As Long
    For i = 0 To UBound(props)
        ' 获取类属性对应的值
        Debug.Print CallByName(p2, props(i), VbGet)
    Next

End Sub