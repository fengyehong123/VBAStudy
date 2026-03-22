Sub MainSub()

    ' 创建一个字典类型的对象
    Dim dict As Object
    ' 使用CreateObject创建的话, 不需要额外的引用, 兼容性好
    ' 如果需要引用的话, 需要在vba编辑器中
    '   工具 → 引用 → 勾选【Microsoft Scripting Runtime】
    '   然后就可以像下面这样创建对象了
    '   Dim dict As New Scripting.Dictionary
    ' 缺点是
    '   如果其他人在运行宏的时候没有勾选, 会报错
    Set dict = CreateObject("Scripting.Dictionary")

    ' 向字典对象添加key和value
    ' 字典中的key是不能重复的, 否则会报错
    dict("Name") = "贾飞天"
    dict("Age") = 23
    dict("Address") = "月球"

    ' 遍历key
    Dim key As Variant
    For Each key In dict.Keys
        Debug.Print key, dict(key)
    Next

    ' 遍历value
    Dim val As Variant
    For Each val In dict.Items
        Debug.Print val
    Next

    ' 判断key是否存在
    If dict.Exists("Name") Then
        Debug.Print "这个key是存在的"
    End If

    ' 获取key的数量
    Debug.Print dict.Count  ' 3

    ' 删除指定的key
    dict.Remove "Name"

    ' 删除全部的key
    dict.RemoveAll
End Sub