Sub MainSub()

    ' 路径分隔符
    Dim delimiter As String: delimiter = Application.PathSeparator

    ' 获取当前用户的桌面路径
    Dim desktopPath As String: desktopPath = Environ("USERPROFILE") & delimiter & "Desktop"
    ' 拼接要读取的文件夹的绝对路径
    Dim folderPath As String: folderPath = desktopPath & delimiter & "Windows技巧"

    ' 获取指定文件夹下的所有文件和文件夹
    Call ListFolderAndFiles(folderPath)

End Sub

' 获取指定文件夹下的所有文件名
Sub ListFolderAndFiles(folderPath As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 若文件夹不存在的话    
    If Not fso.FolderExists(folderPath) Then
        MsgBox "要读取的文件夹并不存在"
        Exit Sub
    End If
    
    ' 获取文件夹对象
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)

    ' 遍历指定文件夹下的所有子文件夹
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        Debug.Print subFolder.Name
    Next
    
    ' 遍历指定文件夹下的所有文件
    Dim file As Object
    For Each file In folder.Files
        Debug.Print file.Name
    Next

End Sub