Sub MainSub()

    ' =====================
    ' 🔷字符串拼接使用 &
    ' =====================
    Const str1 = "Hello,World" & ", 你好, 世界"
    Debug.Print str1

    ' =====================
    ' 🔷字符串长度
    ' =====================
    Debug.Print Len(str1)  ' 20

    ' =====================
    ' 🔷字符串截取
    ' =====================
    ' 从左开始截取
    Debug.Print Left(str1, 3)  ' Hel
    ' 从右开始截取
    Debug.Print Right(str1, 2)  ' 世界
    ' 从指定的位置开始截取
    Debug.Print Mid(str1, 2, 3)  ' ell

    ' =====================
    ' 🔷字符串分隔
    ' =====================
    Dim arr As Variant: arr = Split(str1, ",")
    Debug.Print arr(0)  ' Hello
    Debug.Print arr(1)  ' World

    ' =====================
    ' 🔷字符串去除空格
    ' =====================
    ' 两侧空格
    Trim("  Hello  ")
    ' 左边空格
    LTrim("  Hello")
    ' 右边空格
    RTrim("Hello  ")

    ' =====================
    ' 🔷大小写转换
    ' =====================
    Debug.Print UCase("abc")  ' ABC
    Debug.Print LCase("ABC")  ' abc

    ' =====================
    ' 🔷字符串比较
    ' =====================
    ' 默认比较时区分大小写
    If not "abc" = "ABC" Then
        Debug.Print "不相等"
    End If

    ' =====================
    ' 🔷字符串是否包含
    ' =====================
    If InStr(str1, "世界") > 0 Then
        Debug.Print "字符串中包含世界"
    End If

End Sub

' 提取扩展名
Function GetExt(fileName As String) As String
    GetExt = Mid(fileName, InStrRev(fileName, ".") + 1)
End Function

' 提取文件名
Function GetFileName(path As String) As String
    GetFileName = Mid(path, InStrRev(path, "\") + 1)
End Function