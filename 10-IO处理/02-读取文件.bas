Sub MainSub()

    ' 获取当前用户的桌面路径
    Dim desktopPath As String: desktopPath = Environ("USERPROFILE") & Application.PathSeparator & "Desktop"
    ' 拼接文件的绝对路径
    Dim filePath As String: filePath = desktopPath & Application.PathSeparator & "test.txt"

    ' 读取utf-8格式的文件
    Call ReadUTF8File(filePath)

End Sub

' ===========================================
' 默认的 FileSystemObject（FSO）不支持 UTF-8
' 只能正确处理 ANSI / Unicode（UTF-16）。
' 只能使用 ADODB.Stream 来读取 UTF-8 格式的数据
' ===========================================
Sub ReadUTF8File(filePath As String)

    ' 创建一个 ADODB.Stream 对象
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    ' 文本类型
    Const textType = 2
    ' 文本编码
    Const charsetType = "utf-8"

    ' 🔷读取全部文本
    Dim content As String
    With stream
        ' 文本模式
        .Type = textType
        ' 指定文本的编码类型
        .Charset = charsetType
        ' 打开stream对象
        .Open
        ' 加载要读取的文件路径
        .LoadFromFile filePath
        ' 一口气读取全部内容
        content = .ReadText
        ' 关闭stream对象
        .Close
    End With

    Debug.Print content
    Debug.Print "================"

    ' 🔷按行读取
    Dim line As String
    ' -1 → 全部读取
    ' -2 → 按行读取
    Const lineRead = -2
    With stream
        ' 文本模式
        .Type = textType
        ' 指定文本的编码类型
        .Charset = charsetType
        ' 打开stream对象
        .Open
        ' 加载要读取的文件路径
        .LoadFromFile filePath
        
        Do Until .EOS
            ' 按行进行读取, 读取到的数据中不包含换行符
            line = .ReadText(lineRead)
            Debug.Print line
            Debug.Print "~~~"
        Loop

        ' 关闭stream对象
        .Close
    End With

End Sub