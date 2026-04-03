' 强制要求变量必须要声明
' 如果变量不声明的话, 也能用, 但默认类型是Variant
' vba在使用的时候还需要默认转型
Option Explicit

Sub MainSub()

    ' 定义sheet对象和表对象
    Dim ws As Worksheet
    Dim tblObj As ListObject

    ' 获取指定的Sheet页对象
    Set ws = ThisWorkbook.Worksheets("Sheet3")
    ' 获取指定Sheet页中的指定表对象
    Set tblObj = ws.ListObjects("表3")

    ' 如果当前表格中有数据的话, 就直接清空
    If tblObj.ListRows.count > 0 Then
        tblObj.dataBodyRange.Delete
    End If

    ' 定义一个字典对象
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    ' 向字典对象添加key和value
    dict("MPL_ERR001") = 10
    dict("MPL_ERR002") = 23
    dict("MPL_ERR003") = 26
    dict("QCH_ERR00A") = 52
    dict("QCH_ERR00B") = 45
    dict("QCH_ERR00C") = 28

    ' 遍历字典所用到的变量
    Dim key As Variant
    Dim arr() As String
    Dim row As ListRow
    Dim rowNum As Long

     ' 遍历字典
    For Each key In dict.Keys

        ' 表格新增加一行
        Set row = tblObj.ListRows.Add
        ' 获取当前的行号
        rowNum = tblObj.ListRows.count
        ' 通过下划线分隔字符串
        arr = Split(key, "_")

        ' 为新增的1行的每列都添加数据
        With row.Range
            ' 因为只有1行, 所以 → .Cells(1, 列下标)
            ' 第1列 → No
            .Cells(1, 1) = rowNum
            ' 第2列 → 系统名
            .Cells(1, 2) = arr(0)
            ' 第3列 → 错误码
            .Cells(1, 3) = arr(1)
            ' 第4列 → 值
            .Cells(1, 4) = dict(key)
        End With
    Next

    ' 设置表格的排序属性
    With tblObj.Sort.SortFields
        ' 清除当前表格的排序
        .Clear
        ' === 添加新的排序规则 ===
        '   Key:=tblObj.ListColumns(1).Range → 根据第1列排序
        '   SortOn:=xlSortOnValues → 根据单元格值的值排序
        '   Order:=xlDescending → 倒序（从大到小）
        '   Order:=xlAscending → 正序（从小到大）
        .Add Key:=tblObj.ListColumns(1).Range, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    End With
            
    With tblObj.Sort
        ' 表示表格有表头
        .Header = xlYes
        ' 应用排序
        .Apply
    End With

End Sub