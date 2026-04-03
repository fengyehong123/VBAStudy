Sub MainSub()

    ' 获取当前用户的桌面路径
    Dim desktopPath As String: desktopPath = Environ("USERPROFILE") & Application.PathSeparator & "Desktop"
    ' 拼接文件的绝对路径
    Dim filePath As String: filePath = desktopPath & Application.PathSeparator & "test.txt"

    ' 向桌面写入一个文件, ANSI编码
    Call WriteFileFSO(filePath)
    ' 向桌面写入一个文件, UTF-8编码
    Call WriteUTF8File(filePath)

End Sub

' ========================
' 🔷文件写入
' ========================
' 方式1, FileSystemObject 的方式
' 缺点: 只能使用ANSI的编码, 无法使用UTF-8
Sub WriteFileFSO(filePath As String)

    ' 创建文件系统对象
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 创建文件对象
    Dim ts As Object
    Set ts = fso.CreateTextFile(filePath, True)
    
    ' 向文件中写入内容
    ts.WriteLine "Hello World"
    ts.WriteLine "第二行"
    
    ' 关闭文件对象
    ts.Close

End Sub

' 👍方式2, ADODB.Stream 的方式
' 可以指定文件的编码
' 常量
'   vbLf → LF换行符
'   vbCrLf → CRLF换行符
Sub WriteUTF8File(filePath As String)

    ' ================
    ' 文件删除: 方式1
    ' ================
    ' 若文件存在的话, 则删除
    If Dir(filePath) <> "" Then
        ' kill只能删除文件, 无法删除文件夹
        Kill filePath
    End If

    ' ================
    ' 文件删除: 方式2
    ' ================
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        ' True = 强制删除
        fso.DeleteFile filePath, True
    End If

    ' 创建一个 ADODB.Stream 对象
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    ' 覆盖模式
    Const adSaveCreateOverWrite = 2
    Const textType = 2
    ' 换行flag
    Const adWriteLine = 1

    ' 打开流对象进行处理
    With stream

        ' 流类型为文本
        .Type = textType
        ' 指定utf-8格式的文件
        .Charset = "UTF-8"
        ' 打开流
        .Open

        ' 向文件中写入内容, 由于没有换行符, 所以在一行显示
        .WriteText "你好,世界"
        ' 使用换行符flag
        .WriteText "Hello,World", adWriteLine
        ' 声明使用Windows的CRLF换行符
        .WriteText "测试1,测试2" & vbCrLf
        ' 声明使用Linux的LF换行符
        .WriteText "测试3,测试4" & vbLf
        .WriteText "测试5,测试6" & vbCrLf
        
        ' 覆盖模式保存文件
        .SaveToFile filePath, adSaveCreateOverWrite
        ' 关闭文本流
        .Close

    End With

End Sub