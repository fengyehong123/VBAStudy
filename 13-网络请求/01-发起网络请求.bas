' ========================================
' 做【像浏览器一样请求】 → 用 WinHttp
'   1. 基于 WinHTTP（Windows HTTP Services）
'   2. Windows 提供的 系统级 HTTP 库
'   3. 和浏览器完全独立
' 做【XML/WebService】 → 用 ServerXMLHTTP
'   属于 MSXML（XML 解析库）的一部分
'   本质是为 XML 设计的 HTTP 客户端
' ========================================
Sub SendGetRequest()

    Const url = "https://api.github.com"

    ' 创建http请求对象
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' 发送get请求
    http.Open "GET", url, False
    http.Send

    Debug.Print http.ResponseText

End Sub

Sub SendPostRequest()

    Dim url As String: url = "https://httpbin.org/post"

    ' 1. 创建http请求对象
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' 2. 先设置要发送的请求
    http.Open "POST", url, False

    ' 3. 再设置请求头
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "User-Agent", _
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " & _
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ' 不会自动管理cookie, 所以需要手动添加cookie
    http.setRequestHeader "Cookie", "sessionid=abc123; userid=1001"
    http.setRequestHeader "Accept", "*/*"
    http.setRequestHeader "Accept-Language", "zh-CN,zh;q=0.9"

    ' 发送指定数据
    Dim body As String: body = "{""username"":""admin"",""password"":""123456""}"
    http.Send body

    ' 判断状态码
    If http.Status = 200 Then
        Debug.Print "成功：" & http.ResponseText
    Else
        Debug.Print "失败：" & http.Status
    End If

End Sub