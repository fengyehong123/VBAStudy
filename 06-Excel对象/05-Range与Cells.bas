' ======================================
' Range
'   一块区域（可以是一个或多个单元格）
' Cells
'   一个单元格（通过行列定位）
' ======================================
Sub MainSub()

    ' ==============================
    ' Range 的用法（更人类友好）
    ' ==============================
    ' 定义区域变量
    Dim rng As Range
    Dim cell As Range
    ' 指定当前Excel的指定的Sheet页
    With ThisWorkbook.Worksheets("Sheet1")

        ' 清空当前Sheet页中的所有内容, 包括样式
        .Cells.Clear
        ' 向A1单元格写入内容
        .Range("A1").Value = "Hello World"
        ' 向A2到A5单元格写入内容, 由于是批量赋值, 所以效率比普通的循环要快
        .Range("A2:A5").Value = 100
        
        ' 创建一个区域对象
        Set rng = .Range("D1:H1")
        rng.Value = "测试内容"

        ' 遍历Range区域, 每一个元素都是一个Cell
        For Each cell In rng
            Debug.Print cell.Value
        Next cell

    End With
    Debug.Print "==================="

    ' ==============================
    ' Cells 的用法（更程序化）
    '   Cells(行, 列)
    ' ==============================
    Dim i As Integer
    ' 指定当前Excel的指定的Sheet页
    With ThisWorkbook.Worksheets("Sheet2")

        ' 清空当前Sheet页中的所有内容, 包括样式
        .Cells.Clear
        ' 向第1行, 第1列写入内容 → A1
        .Cells(1, 1) = "你好"
        ' 向第2行, 第3列写入内容 → C2
        .Cells(2, 3) = "我好"

        ' 循环写数据: 向A2到A6写入数据
        For i = 2 To 6
            .Cells(i, 1).Value = i - 1
        Next i

        ' 等价于 Range("A9:C10")
        .Range(.Cells(9, 1), .Cells(10, 3)).Value = "默认值"

    End With

End Sub