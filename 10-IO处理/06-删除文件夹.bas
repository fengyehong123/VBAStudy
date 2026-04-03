Sub MainSub()

    ' 路径分隔符
    Dim delimiter As String: delimiter = Application.PathSeparator

    ' 获取当前用户的桌面路径
    Dim desktopPath As String: desktopPath = Environ("USERPROFILE") & delimiter & "Desktop"
    ' 要删除的文件夹
    Dim folderPath As String: folderPath = desktopPath & delimiter & "TestFolder"

    ' 删除指定的文件夹
    Call DeleteFolderFSO(folderPath)

End Sub

Sub DeleteFolderFSO(folderPath As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        MsgBox "要删除的文件夹并不存在"
        Exit Sub
    End If

    ' True = 强制删除（包括里面的文件）
    fso.DeleteFolder folderPath, True
End Sub