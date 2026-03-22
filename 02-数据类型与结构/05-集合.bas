' ================================================
' Collection 是 VBA 自带的一个【集合类】，特点是：
'   1. 按顺序存储数据（有序）
'   2. 可以用索引访问（类似数组）
'   3. 也可以用 Key（可选）
' ================================================
Sub MainSub()

    ' 创建一个集合
    Dim col_1 As New Collection

    ' 向集合中添加元素
    col_1.Add "苹果"
    col_1.Add "香蕉"
    col_1.Add "橙子"

    ' 使用索引读取元素
    ' 💥注意: 从下标1开始
    Debug.Print col_1(1)  ' 苹果

    ' 获取集合的数量
    Debug.Print col_1.Count

    ' 使用索引删除集合中的元素
    ' 删除元素的时候只能使用索引, 不能使用key
    col_1.Remove 1
    Debug.Print col_1.Count
    Debug.Print "-----------------------"

    ' 集合中元素的遍历
    Dim item As Variant
    For Each item In col_1
        Debug.Print item
    Next
    Debug.Print "-----------------------"

    ' 使用for循环进行遍历
    Dim i As Long
    For i = 1 To col_1.Count
        Debug.Print col_1(i)
    Next
    Debug.Print "-----------------------"

    ' 插入元素到指定位置之前
    ' 插入到第1个位置之前
    col_1.Add "西瓜", , 1
    For Each item In col_1
        Debug.Print item
    Next
    ' 西瓜
    ' 香蕉
    ' 橙子

    ' 再创建一个集合
    Dim col_2 As New Collection
    ' 还可以使用Key添加, 类似 Dictionary
    col_2.Add "张三", "A001"
    col_2.Add "李四", "A002"

    ' key不能重复, 否则会报错
    ' col_2.Add "王五", "A002"
    ' vba中并没有判断集合中的key是否存在的方法
    
    ' 使用key来获取数据
    Debug.Print col_2("A001")
    Debug.Print "-----------------------"

    ' 使用可以初始化集合的函数
    Dim col_3 As Collection
    Set col_3 = NewCollection("A", "B", "C")
    For Each item In col_3
        Debug.Print item
    Next

End Sub

' 封装一个用来可以初始化集合的函数
Function NewCollection(ParamArray items() As Variant) As Collection

    Dim col As New Collection
    Dim i As Long
    
    For i = LBound(items) To UBound(items)
        col.Add items(i)
    Next
    
    Set NewCollection = col
End Function