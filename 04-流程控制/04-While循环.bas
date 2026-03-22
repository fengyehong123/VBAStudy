Sub MainSub()

    Dim i As Long

    ' 普通的While循环
    Do While i <= 5

        If i = 5 Then
            ' 一定要记得 + 1
            i = i + 1
            ' 跳出此次循环
            GoTo ContinueLoop
        End If

        If i = 6 Then
            ' 终止此次循环
            Exit Do
        End If

        Debug.Print i
        i = i + 1

    ContinueLoop:
    Loop

    ' Until循环
    Do Until i > 5
        Debug.Print i
        i = i + 1
    Loop

    ' 先执行再判断, 至少执行一次
    Do
        Debug.Print i
        i = i + 1
    Loop While i <= 5
    
End Sub