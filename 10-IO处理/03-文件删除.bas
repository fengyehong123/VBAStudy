Sub MainSub()

   ' 获取当前用户的桌面路径
    Dim desktopPath As String: desktopPath = Environ("USERPROFILE") & "\Desktop"
    ' 拼接文件的绝对路径
    Dim filePath As String: filePath = desktopPath & Application.PathSeparator & "test.txt"

    DeleteFileIfExists(filePath)

End Sub

' 若文件存在则删除
Function DeleteFileIfExists(filePath As String) As Boolean

    On Error Resume Next
    
    If Dir(filePath) <> "" Then
        Kill filePath
        DeleteFileIfExists = True
    Else
        DeleteFileIfExists = False
    End If

End Function