Sub MainSub()

    Dim i As Long

    ' 🔷最基本的for循环
    For i = 1 To 5
        Debug.Print i
    Next i

    ' 🔷指定步长的for循环
    For i = 1 To 10 Step 2
        Debug.Print i
    Next i

    ' 🔷反向for循环
    For i = 5 To 1 Step -1
        Debug.Print i
    Next i

    ' 🔷满足条件终止for循环
    For i = 1 To 10

        If i = 5 Then
            Exit For
        End If

        Debug.Print i
    Next i

    ' 🔷满足条件, 跳过本次for循环
    For i = 1 To 10

        If i = 5 Then
            ' 💥vba中没有continue关键字, 只能使用 GoTo 进行跳过
            GoTo NextHandler
        End If

        Debug.Print i
    
    NextHandler:
    Next i
    Debug.Print "-----------------------"

    ' 创建一个集合
    Dim col As New Collection
    ' 向集合中添加元素
    col.Add "苹果"
    col.Add "香蕉"
    col.Add "橙子"

    ' 🔷使用 For Each 来遍历集合
    Dim item As Variant
    For Each item In col

        If item = "苹果" Then
            ' 🔷使用 GoTo 跳出本次循环
            GoTo NextItem
        End If

        If item = "橙子" Then
            ' 🔷终止循环
            Exit For
        End If

        Debug.Print item

    NextItem:
    Next item

End Sub