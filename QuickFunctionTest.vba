' 简单的函数存在性测试
Sub TestGetElapsedTime()
    Debug.Print "测试 GetElapsedTime 函数..."
    
    ' 获取开始时间
    Dim startTime As Long
    startTime = GetTickCount()
    
    ' 等待一点时间
    Dim i As Long
    For i = 1 To 1000
        ' 简单循环
    Next i
    
    ' 计算耗时
    Dim elapsed As Double
    elapsed = GetElapsedTime(startTime)
    
    Debug.Print "开始时间: " & startTime
    Debug.Print "耗时: " & elapsed & " 秒"
    Debug.Print "✅ GetElapsedTime 函数测试成功！"
End Sub

' 测试 StartTimer 函数
Sub TestStartTimer()
    Debug.Print "测试 StartTimer 函数..."
    
    Dim timer As Long
    timer = StartTimer()
    
    Debug.Print "计时器值: " & timer
    Debug.Print "✅ StartTimer 函数测试成功！"
End Sub

' 测试 ClearProgress 函数
Sub TestClearProgress()
    Debug.Print "测试 ClearProgress 函数..."
    
    Application.StatusBar = "测试状态栏"
    Debug.Print "设置状态栏: 测试状态栏"
    
    ClearProgress
    Debug.Print "✅ ClearProgress 函数测试成功！"
End Sub

' 运行所有测试
Sub RunQuickTests()
    Debug.Print "==============================="
    Debug.Print "快速函数测试"
    Debug.Print "==============================="
    
    TestStartTimer
    TestGetElapsedTime
    TestClearProgress
    
    Debug.Print "==============================="
    Debug.Print "✅ 所有快速测试完成！"
    Debug.Print "==============================="
End Sub
