Sub MainSub()
    ' 获取文件的弹窗
    Call FileDialogSub
    ' 获取文件夹的弹窗
    Call FolderDialogSub
End Sub

' 文件获取弹窗
private Sub FileDialogSub()

    ' 文件路径
    Dim filePath As String

    ' ============================
    ' 🔷用户选择指定文件
    ' ============================
    ' 方式1
    filePath = Application.GetOpenFilename("Excel文件 (*.xlsx), *.xlsx")
    If filePath = "False" Then
        MsgBox "你取消了操作"
        Exit Sub
    End If

    Debug.Print "用户选择的文件为: " & filePath
    
    ' 方式2
    With Application.FileDialog(msoFileDialogFilePicker)

        ' 设置标题
        .Title = "请选择文件"
        ' 设置打开对话框时的默认路径
        .InitialFileName = "C:\Users\"

        ' 💥限制文件类型
        .Filters.Clear
        .Filters.Add "Excel 文件", "*.xlsx; *.xls"
        .Filters.Add "所有文件", "*.*"
        
        ' 允许多选文件
        '.AllowMultiSelect = True

        ' 如果用户未选择文件夹则退出程序
        If .Show Then 
            ' 只能选一个文件的情况下
            filePath = .SelectedItems(1)
            ' 多选的情况下, 遍历被选中的文件
            ' For Each filePath In .SelectedItems
            '     Debug.Print filePath
            ' Next 
        Else
            Exit Sub
        End If

    End With 
    Debug.Print "用户选择的文件为: " & filePath

End Sub

' 文件夹获取弹窗
private Sub FolderDialogSub

    ' 文件夹路径
    Dim folderPath As String

    ' 🔷用户选择指定文件夹
    With Application.FileDialog(msoFileDialogFolderPicker)
        ' 设置标题
        .Title = "请选择指定文件夹。"
        If .Show Then
            ' 获取文件夹路径
            folderPath = .SelectedItems(1)
        Else
            ' 如果没有选择任何文件夹的话, 则终止当前Sub
            Exit Sub
        End If
    End With
    Debug.Print folderPath

End Sub