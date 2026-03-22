' 定义一个枚举
' 如果不写在模块中, 要和Sub写在一个文件中的话
'   只能放在顶部
Enum Gender
    ' 对应的只能是数字, 不能是字符串
    Male = 1
    Female = 2
End Enum

Enum Colors
  Red = 1
  Green = 2
  Blue = 3
End Enum

' 定义一个枚举转换字符串的函数
Function GenderToString(g As Gender) As String
    Select Case g
        Case Male: GenderToString = "男"
        Case Female: GenderToString = "女"
        Case Else: GenderToString = "未知"
    End Select
End Function

Sub MainSub()

    ' 创建一个枚举对象
    Dim gener As Gender: gener = Male
    ' 枚举转换字符串
    Dim str As String: str = GenderToString(gener)
    Debug.Print str  ' 男

    ' 提高可读性
    Debug.Print Colors.Red  ' 1

    ' 如果性别不是Female的话
    If not gener = Female Then
        Debug.Print "性别不是Female"
    End If
End Sub