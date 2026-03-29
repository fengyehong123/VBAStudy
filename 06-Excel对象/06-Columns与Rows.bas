Sub MainSub()
    
    ' ========================
    ' Rows
    '   表示工作表中的一整行
    ' Columns
    '   表示工作表中的一整列
    ' ========================

    ' 调整A列的宽度
    Columns("A").ColumnWidth = 20
    ' 调整第1行的高度
    Rows(1).RowHeight = 30

    ' ==============================
    ' lastRow 的用法
    '   找到表格的最后一行
    ' ==============================
    Dim lastRow As Long
    With ThisWorkbook.Worksheets("Sheet3")

        ' 找到A列的最后一行
        '   .End(xlUp) → 向上找第一个有值的单元格
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

        ' 打印A1到A最后一行的值
        For i = 1 To lastRow
            Debug.Print .Cells(i, 1).Value
        Next i

    End With

End Sub