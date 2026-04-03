Option Explicit

Sub MainSub()

    ' 定义sheet对象和表对象
    Dim ws As Worksheet
    Dim tblObj As ListObject

    ' 获取指定的Sheet页对象
    Set ws = ThisWorkbook.Worksheets("Sheet3")
    ' 获取指定Sheet页中的指定表对象
    Set tblObj = ws.ListObjects("表3")

    ' 如果当前表格中没有数据的话, 提示用户
    If tblObj.ListRows.count <= 0 Then
        MsgBox "表格中没有数据, 请确认..."
        Exit Sub
    End If

    ' 表格相关变量
    Dim systemNameCell As Range
    Dim systemNameRange As Range

    ' 🔷根据表格的列名获取当前列的数据(不会包含表头的数据)
    Set systemNameRange = tblObj.ListColumns("系统名").DataBodyRange
    If systemNameRange Is Nothing Then
        MsgBox "当前列, 没有数据..."
        Exit Sub
    End If

    ' 遍历表格中指定列的数据
    For Each systemNameCell In systemNameRange
        Debug.Print systemNameCell.Value
    Next
    Debug.Print "============================"

    ' 🔷根据表格的列索引来获取当前列的数据并遍历
    For Each systemNameCell In tblObj.ListColumns(2).DataBodyRange
        Debug.Print systemNameCell.Value
    Next
    Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~"

    ' ==============================
    ' 👍推荐写法: 
    '   直接读取当前列的数组
    '   然后遍历, 效率会更高
    ' 
    ' 如果一个个遍历单元格的话
    ' 当数据多的时候, 会影响效率
    ' ==============================
    Set systemNameRange = tblObj.ListColumns(2).DataBodyRange
    If systemNameRange Is Nothing Then
        MsgBox "当前列, 没有数据..."
        Exit Sub
    End If
    
    ' 🔷获取包含当前列中的所有的数据的数组
    Dim i As Long
    Dim dataArr As Variant: dataArr = systemNameRange.Value

    For i = 1 To UBound(dataArr, 1)
        Debug.Print dataArr(i, 1)
    Next

End Sub