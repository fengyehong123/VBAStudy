' ===================================
' VBA里的正则表达式, 主要依赖于
'   VBScript.RegExp 对象
' 
' 该正则对象并不支持高级正则语法, 例如
'   (?<=...) 后行断言
'   (?<!...) 负向断言
'
' 需要注意的是
'   默认不是全局匹配, 需要手动开启
' ===================================
Sub MainSub()

    ' 创建一个正则对象
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")

    ' 为正则对象设置对应的属性
    Const regStr = "\d+"
    With reg
        ' 正则表达式
        .Pattern = regStr
        ' 是否匹配全部
        .Global = True        
        ' 是否忽略大小写
        .IgnoreCase = True    
    End With

    ' 🔷Test测试是否匹配
    Dim result1 As Boolean: result1 = reg.Test("abc123")
    Debug.Print result1  ' True

    ' 🔷Execute获取匹配内容
    Dim matches As Object
    Set matches = reg.Execute("abc123def456")

    Dim item As Object
    For Each item In matches
        Debug.Print item.Value
        ' 123
        ' 456
    Next

    ' 🔷Replace替换
    Dim result2 As String: result2 = reg.Replace("abc123def", "#")
    Debug.Print result2  ' abc#def

    ' 🔷分组处理
    reg.Pattern = "(\d{4})-(\d{2})-(\d{2})"
    Set matches = reg.Execute("2026-03-29")
    Debug.Print matches(0).SubMatches(0)  ' 2026
    Debug.Print matches(0).SubMatches(1)  ' 03
    Debug.Print matches(0).SubMatches(2)  ' 29

End Sub