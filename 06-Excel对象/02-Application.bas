Sub MainSub()

    ' 获取当前Application的名称
    Debug.Print Application.Name  ' Microsoft Excel

    ' 关闭提示框 → 删除工作表不弹出确认框
    Application.DisplayAlerts = False

    ' 隐藏与显示Excel
    Application.Visible = False
    Application.Visible = True

    ' 获取路径
    Debug.Print Application.Path  ' C:\Program Files (x86)\Microsoft Office\Root\Office16
    Debug.Print Application.StartupPath  ' C:\Users\用户名\AppData\Roaming\Microsoft\Excel\XLSTART

    ' ====================== 性能优化 ======================
    ' 🔷业务处理之前先关闭屏幕刷新和自动计算
    ' 关闭屏幕刷新
    Application.ScreenUpdating = False
    ' 关闭自动计算
    Application.Calculation = xlCalculationManual

    ' vba中的业务计算1
    ' vba中的业务计算2
    ' vba中的业务计算3

    ' 🔷业务处理之后再开启屏幕刷新和自动计算
    ' 开启自动计算
    Application.Calculation = xlCalculationAutomatic
    ' 恢复屏幕刷新
    Application.ScreenUpdating = True
    ' ====================== 性能优化 ======================

    ' 路径分隔符
    Debug.Print Application.PathSeparator  ' \
    
End Sub