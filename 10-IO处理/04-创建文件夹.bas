Sub MainSub()

    Dim delimiter As String: delimiter = Application.PathSeparator

    ' 获取当前用户的桌面路径
    Dim desktopPath As String: desktopPath = Environ("USERPROFILE") & delimiter & "Desktop"
    ' 拼接要创建的文件夹的绝对路径
    Dim folderPath As String: folderPath = desktopPath & delimiter & "TestFolder" & delimiter & "01"

    ' 创建文件夹
    Call CreateFolderFSO(folderPath)

End Sub

' 通过 FileSystemObject 的方式创建文件夹
Sub CreateFolderFSO(folderPath As String)

    ' 路径分隔符
    Dim delimiter As String: delimiter = Application.PathSeparator

    ' 根据路径分隔符将文件夹路径分割为数组
    Dim parts() As String: parts = Split(folderPath, delimiter)
    ' 获取一部分路径
    Dim currentPath As String: currentPath = parts(0)

    ' ======================================
    ' FileSystemObject 并不支持递归创建文件夹
    ' 需要每一层都手动创建
    ' ======================================
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 递归手动创建每一层文件夹
    Dim i As Long
    For i = 1 To UBound(parts)

        currentPath = currentPath & delimiter & parts(i)
        
        ' 若文件夹不存在的话    
        If Not fso.FolderExists(currentPath) Then
            ' 创建文件夹
            fso.CreateFolder currentPath
        End If
    Next i

End Sub