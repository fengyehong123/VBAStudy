Sub MainSub()

    Dim score As Long: score = 100

    If score >= 90 Then
        Debug.Print "优秀"
    ElseIf score >= 60 Then
        Debug.Print "及格"
    Else
        Debug.Print "不及格"
    End If

End Sub