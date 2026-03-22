' ========================================
' 🔷Excel VBA 时最常用的对象数据类型
'   Object	    通用对象
'   Workbook    工作簿
'   Worksheet   工作表
'   Range       单元格
'   Cells	    单元格集合
' ========================================
Sub MainSub()

    Dim ws As Worksheet
    ' 对象赋值的时候必须使用Set
    Set ws = ThisWorkbook.Sheets("Sheet1")

    Dim r As Range
    Set r = ws.Range("A1")
    
End Sub