Sub MainSub()

    ' ====================
    ' 🔷刷新整个Excel
    ' ====================
    ' 方式一: 刷新当前打开的整个Excel, 简单粗暴, 最省事
    '   PowerQuery创建的表格
    '   数据连接
    '   数据透视表
    ' 会被一键刷新
    'ThisWorkbook.RefreshAll

    ' ====================
    ' 🔷刷新连接对象
    ' ====================
    ' 方式二: 刷新PowerQuery的连接对象
    Dim connectObj As WorkbookConnection
    ' 获取当前Excel文件中的所有连接对象
    For Each connectObj In ThisWorkbook.Connections

        ' 打印连接名字
        Debug.Print connectObj.Name ' 查询 - 表1
        ' 打印连接的类型
        Debug.Print connectObj.Type

        ' Power Query 一般名字是：查询 - xxx
        If InStr(connectObj.Name, "查询") > 0 Then           
            ' 刷新连接
            connectObj.Refresh
        End If
    Next
    ' 根据连接名找到连接Powerquery的连接对象, 然后刷新
    ThisWorkbook.Connections("查询 - 表1").Refresh

    ' ====================
    ' 🔷刷新表格对象
    ' ====================
    ' 方式三: 刷新指定的表格对象
    Dim currentWs As Worksheet
    Dim tbl1 As ListObject

    ' 获取当前打开的Excel的指定的Sheet页
    Set currentWs = ThisWorkbook.Worksheets("Sheet2")
    ' 获取指定Sheet页中的指定表格
    Set tbl1 = currentWs.ListObjects("表1_2")

    ' 获取当前表格的类型
    ' 0 普通表格
    ' 1 外部数据
    ' 2 查询
    Debug.Print tbl1.SourceType
    
    ' 如果当前表格对象存在就刷新
    If Not tbl1.QueryTable Is Nothing Then
        tbl1.QueryTable.Refresh BackgroundQuery:=False
    End If
    
End Sub