Sub MainSub()

    Dim score As Long: score = 100

    Select Case score
        Case Is >= 90
            Debug.Print "优秀"
        Case Is >= 60
            Debug.Print "及格"
        Case Else
            Debug.Print "不及格"
    End Select

    Select Case score
        Case 100
            Debug.Print "满分"
        Case 90 To 99
            Debug.Print "优秀"
        Case Else
            Debug.Print "其他"
    End Select
    
End Sub