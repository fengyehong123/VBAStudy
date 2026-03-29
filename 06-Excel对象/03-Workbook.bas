Sub MainSub()

    Dim wb As Workbook
    Dim ws As Worksheet

    ' ThisWorkbook → 当前vba代码所在Excel文件
    ' 获取指定的Sheet对象
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    ' ====================================================
    ' 🔷获取已经被打开的指定文件
    '   如果文件没有被打开的话, 会报下标越界错误
    '   此时可以使用【On Error】包裹, 保证程序可以继续执行
    ' ====================================================
    On Error Resume Next
    Set wb = Workbooks("测试.xlsx")
    On Error GoTo 0

    If wb Is Nothing Then
        Debug.Print "文件还没有被打开呢..."
    End If

    ' =====================
    ' 遍历所有已经打开的文件
    ' =====================
    Dim found As Boolean
    For Each wb In Workbooks
        If wb.Name = "测试.xlsx" Then
            found = True
            Exit For
        End If
    Next

    If Not found Then
        Debug.Print "文件还没有被打开呢..."
    End If

    ' 获取当前打开的Excel的同一级目录下的其他Excel文件
    Const otherFileName = "新建 Microsoft Excel 工作表.xlsm"
    Dim otherFileFullPath As String: otherFileFullPath = ThisWorkbook.Path & "\" & otherFileName

    ' 如果当前文件不存在的话
    If Dir(otherFileFullPath) = "" Then
        MsgBox otherFileFullPath & "并不存在, 请重新确认!", vbExclamation
        Exit Sub
    End If
    
    ' 关闭屏幕刷新
    Application.ScreenUpdating = False

    ' 以只读方式打开指定的Excel文件
    Dim otherWorkbook As Workbook
    Set otherWorkbook = Workbooks.Open(otherFileFullPath, ReadOnly:=True)

    ' 遍历所有的Sheet
    Dim otherWorksheet As Worksheet
    For Each otherWorksheet In otherWorkbook.Worksheets
        Debug.Print otherWorksheet.Name
    Next

    ' ===================================
    ' 不是只读模式的话, 保存和关闭一起使用
    ' ===================================
    ' 保存
    ' otherWorkbook.Save
    ' 关闭
    ' otherWorkbook.Close

    ' 另存为
    otherWorkbook.SaveAs otherWorkbook.Path & "\" & "另存为文件.xlsm"

    ' 不保存, 关闭打开的Excel
    otherWorkbook.Close SaveChanges:=False

    ' 恢复屏幕刷新
    Application.ScreenUpdating = True

End Sub