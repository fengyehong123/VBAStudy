Sub MainSub()

    ' 获取桌面路径
    Debug.Print GetDesktopPath1()
    Debug.Print GetDesktopPath2()

End Sub

' ===========================
' 🔷获取桌面文件路径
' ===========================
' 方式一
Function GetDesktopPath1() As String
    ' 获取 USERPROFILE 的环境变量
    GetDesktopPath1 = Environ("USERPROFILE") & "\Desktop"
End Function

' 方式二
Function GetDesktopPath2() As String
    Dim obj As Object
    Set obj = CreateObject("WScript.Shell")
    ' 获取桌面文件夹路径
    GetDesktopPath2 = obj.SpecialFolders("Desktop")
End Function
