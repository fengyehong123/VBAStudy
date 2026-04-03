' =================================
' 图标类型
' vbInformation ℹ️ 信息提示（常用）
' vbExclamation ⚠️ 警告
' vbCritical    ❌ 错误
' vbQuestion    ❓ 提问
' =================================

' =================================================
' 按钮类型
' vbOKOnly              OK
' vbOKCancel            OK / Cancel
' vbYesNo               Yes / No
' vbYesNoCancel         Yes / No / Cancel
' vbRetryCancel         Retry / Cancel
' vbAbortRetryIgnore    Abort / Retry / Ignore
' =================================================
Sub MainSub()

    ' 最基本的消息弹窗
    MsgBox "处理完成!", vbInformation

    ' ==========================================
    ' 🔷带按钮 + 获取用户选择的弹窗
    ' 常用参数
    '   vbYesNo vbOKCancel
    '   vbInformation vbExclamation vbCritical
    ' ==========================================
    Dim result As VbMsgBoxResult: result = MsgBox("是否继续?", vbYesNo + vbQuestion, "确认")
    If result = vbYes Then
        MsgBox "你点了 Yes"
    Else
        MsgBox "你点了 No"
    End If

End Sub