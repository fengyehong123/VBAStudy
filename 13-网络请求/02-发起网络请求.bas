' 发起get请求
Sub SendGetRequest()
    
    Const url = "https://api.github.com"

    ' 创建http请求对象
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.Open "GET", url, False
    http.Send

    Debug.Print http.responseText

End Sub

' 发起post请求
Sub SendPostRequest()

    Const url = "https://httpbin.org/post"

    ' 1. 创建http请求对象
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' 2. 先设置要发送的请求
    http.Open "POST", url, False

    ' 3. 再设置请求头
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "User-Agent", _
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " & _
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ' MSXML2.ServerXMLHTTP.6.0不会自动管理cookie, 所以需要手动添加cookie
    http.setRequestHeader "Cookie", "sessionid=abc123; userid=1001"
    http.setRequestHeader "Accept", "*/*"
    http.setRequestHeader "Accept-Language", "zh-CN,zh;q=0.9"

    ' 4. 然后再发送数据
    http.Send "{""name"":""test"",""age"":18}"

    ' 判断状态码
    If http.Status = 200 Then
        ' 打印返回值
        Debug.Print http.responseText
    Else
        ' 打印状态码
        Debug.Print "错误：" & http.Status
    End If

End Sub