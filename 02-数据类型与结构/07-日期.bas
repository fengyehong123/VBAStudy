Sub MainSub()

    ' =======================================
    ' 🔷定义日期
    '   # # 是 VBA 的日期字面量
    '   日期格式为yyyy-mm-dd
    ' =======================================
    Dim d1 As Date: d1 = #2026-03-21#
    Debug.Print d1  ' 2026/3/21
    ' 日期
    Dim d2 As Date: d2 = Date
    Debug.Print d2  ' 2026/3/21
    ' 日期和时间
    Dim d3 As Date: d3 = Now
    Debug.Print d3  ' 2026/3/21 21:44:47
    ' ✅推荐使用DateSerial, 不会受到地区的影响
    Dim d4 As Date: d4 = DateSerial(2025, 3, 21)
    Debug.Print d4  ' 2025/3/21

    ' =======================================
    ' 🔷日期的本质
    '   VBA 的 Date 本质是一个 Double 数字：
    '       整数部分 → 日期（天数）
    '       小数部分 → 时间
    ' =======================================
    Debug.Print CDbl(#2026-03-21#)  ' 46102
    ' Double数字转换日期
    Debug.Print CDate(46102)  ' 2026/3/21
    Debug.Print "============================"

    ' 字符串转换日期
    Dim d5 As Date: d5 = CDate("2025-03-21")

    ' =======================================
    ' 🔷日期的格式化
    ' =======================================
    Dim str1 As String: s = Format(d5, "yyyy/mm/dd")
    Debug.Print str1

    ' 当前时刻
    Dim d6 As Date: d6 = Now
    Debug.Print Format(d6, "yyyy-mm-dd")  ' 2026-03-21
    Debug.Print Format(d6, "yyyy年mm月dd日")  ' 2026年03月21日
    ' 注意: 分钟是 nn, 不是 mm
    Debug.Print Format(d6, "hh:nn:ss")  ' 21:57:59

    ' 分别获取年月日
    Debug.Print Year(d6)  ' 2026
    Debug.Print Month(d6)  ' 3
    Debug.Print Day(d6)  ' 21

    ' =======================================
    ' 🔷日期的计算
    ' =======================================
    ' 明天
    Debug.Print d6 + 1  ' 2026/3/22 22:08:58
    ' 昨天
    Debug.Print d6 - 1  ' 2026/3/20 22:08:58
    ' +7 天
    Debug.Print DateAdd("d", 7, d6)  ' 2026/3/28 22:10:43
    ' +1个月
    Debug.Print DateAdd("m", 1, d6)  ' 2026/4/21 22:10:43
    ' +1年
    Debug.Print DateAdd("yyyy", 1, d6)  ' 2027/3/21 22:10:43

    ' =======================================
    ' 🔷日期的差值
    ' =======================================
    Dim diff As Long
    diff = DateDiff("d", #2025-03-01#, #2025-03-21#)
    Debug.Print diff  ' 20

End Sub