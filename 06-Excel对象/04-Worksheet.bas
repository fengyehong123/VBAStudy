Sub MainSub()

    Dim ws As Worksheet

    ' 根据Sheet页的名字获取Sheet对象
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    ' 根据索引获取Sheet对象
    ' 注意:
    '   索引从 1 开始
    '   顺序变化会影响结果
    Set ws = ThisWorkbook.Worksheets(1)

    ' =================================================
    ' ⭐高级用法 → CodeName
    '   1. 不受 Sheet 名称修改影响
    '   2. 最稳定
    '
    ' 在VBA里, 每个工作表其实有两个名字：
    '   Name（表名）→ Sheet1 → Excel 标签页看到的名字
    '   CodeName（代码名）→ Sheet1 → VBA 里用的名字
    '
    ' 🔷在此处的 Sheet1 是一个【预定义对象】, 不是普通变量
    '   而是 VBA 自动生成的全局 Worksheet 对象（CodeName）
    '   除非在VBA中的属性窗口中手动修改, 否则是一直存在的
    ' =================================================
    Sheet1.Range("A1").Value = "Hello"

    ' 新增加一个Sheet页
    Dim addWorkSheet As Worksheet
    Const addWorkSheetName = "测试Sheet页"
    Const renameWorkSheetName = "重命名测试Sheet页"

    ' 判断要增加的Sheet页是否存在
    If SheetExists(addWorkSheetName) Then
        MsgBox "要增加的名称为" & addWorkSheetName & "的Sheet页已存在!"
        Exit Sub
    End If

    With ThisWorkbook
        ' 增加一个Sheet页, 并命名
        Set addWorkSheet = .Worksheets.Add
        addWorkSheet.Name = addWorkSheetName

        ' 给刚创建的Sheet页重命名
        .Worksheets(addWorkSheetName).Name = renameWorkSheetName

        ' 删除Sheet页, 关闭提示, 否则会弹窗框让用户确认是否要删除
        Application.DisplayAlerts = False
        ' 为防止删除报错, 永远不会弹出让用户确认的弹窗的问题, 使用 GoTo 
        ' 当删除报错时, 直接跳转到【CleanUp:】, 然后恢复弹窗提示
        On Error GoTo CleanUp
        .Worksheets(renameWorkSheetName).Delete
        CleanUp:
        Application.DisplayAlerts = True

        ' 再增加一个Sheet页
        .Worksheets.Add.Name = addWorkSheetName & "_1"
        ' 将增加的Sheet页隐藏
        .Worksheets(addWorkSheetName & "_1").Visible = False
    End With

    ' 遍历所有的Sheet页
    Dim itemWorkSheet As Worksheet
    For Each itemWorkSheet In ThisWorkbook.Worksheets
        If itemWorkSheet.Name Like "*数据*" Then
            itemWorkSheet.Range("A1").Value = "OK"
        End If
    Next

End Sub

' 判断指定名字的Sheet页是否存在
Function SheetExists(name As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(name) Is Nothing
    On Error GoTo 0
End Function