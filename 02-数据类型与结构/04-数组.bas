Enum Gender
    ' 对应的只能是数字, 不能是字符串
    Male = 1
    Female = 2
End Enum

Function GenderToString(g As Gender) As String

    ' 使用Static关键字初始化一个数组
    ' 一旦数组被赋值之后, 就会被缓存, 不会每次都被初始化, 有助于提升效率
    Static arr As Variant
    ' 判断数组是否为空
    If IsEmpty(arr) Then arr = Array("未知", "男", "女")

    If g >= 0 And g <= UBound(arr) Then
        GenderToString = arr(g)
    Else
        GenderToString = "未知"
    End If

End Function

Sub MainSub()

    ' ============================================
    ' 使用Array初始化一个数组
    ' Array初始化数组的元素的类型是 Variant
    ' Variant 数组（最灵活）
    '   不需要声明大小
    '   类型自动适配
    '   常用于快速初始化
    ' ============================================
    Dim props As Variant: props = Array("Name", "Age")
    
    ' 遍历数组
    Dim i As Integer
    For i = LBound(props) To UBound(props)
        Debug.Print props(i)
        ' Name
        ' Age
    Next i
    Debug.Print "-----------------------"

    ' 遍历数组
    Dim item As Variant
    For Each item In props
        Debug.Print item
        ' Name
        ' Age
    Next
    Debug.Print "-----------------------"

    ' 清空数组
    Erase props

    ' 声明固定长度的数组
    Dim arr1(1 To 5) As Integer
    ' 这种写法声明的数组
    '   元素个数：5个
    '   下标范围：1 ~ 5
    arr1(1) = 10
    arr1(2) = 20
    arr1(3) = 30
    arr1(4) = 40
    arr1(5) = 50
    ' 下标起始和终了位置
    Debug.Print LBound(arr1)  ' 1
    Debug.Print UBound(arr1)  ' 5
    Debug.Print "-----------------------"

    Dim arr2(5) As Integer
    ' 这种写法声明的数组
    '   元素个数：6个
    '   下标范围：0 ~ 5（默认从0开始）
    arr2(0) = 10
    arr2(1) = 20
    arr2(2) = 30
    arr2(3) = 40
    arr2(4) = 50
    arr2(5) = 60
    ' 下标起始和终了位置
    Debug.Print LBound(arr2)  ' 0
    Debug.Print UBound(arr2)  ' 5
    Debug.Print "-----------------------"

    ' 动态数组
    Dim arr3() As Integer
    ' 动态数组必须用 ReDim 才能使用
    ReDim arr3(1 To 5)

    ' 创建二维数组
    Dim arr4(1 To 3, 1 To 2) As Integer
    ' 向二维数组中添加元素
    arr4(1, 1) = 10
    arr4(1, 2) = 20
    arr4(2, 1) = 30
    arr4(2, 2) = 40
    arr4(3, 1) = 50
    arr4(3, 2) = 60

    ' 遍历二维数组
    Dim z As Integer, j As Integer
    For z = 1 To 3
        For j = 1 To 2
            Debug.Print arr4(z, j)
            ' 10 
            ' 20 
            ' 30 
            ' 40 
            ' 50 
            ' 60 
        Next j
    Next z

    ' 判断是否是数组
    Debug.Print IsArray(arr4)  ' True
    
End Sub