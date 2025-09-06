Sub BatchCopyTrackingNumbers()
    '批量复制快递单号工具（支持续复制）
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim trackingNumbers As String
    Dim i As Long
    Dim count As Integer
    Dim copyCount As Integer
    Dim userInput As String
    Dim startRow As Long
    Dim lastCopiedRow As Long
    
    Set ws = ActiveSheet
    
    '获取用户输入的复制数量
    userInput = InputBox("请输入要复制的快递单号数量：" & vbCrLf & _
                        "建议范围：1-20个", "批量复制快递单号", "5")
    
    '验证用户输入
    If userInput = "" Then Exit Sub
    
    If IsNumeric(userInput) Then
        copyCount = CInt(userInput)
        If copyCount < 1 Then
            MsgBox "请输入大于0的数字！"
            Exit Sub
        End If
    Else
        MsgBox "请输入有效的数字！"
        Exit Sub
    End If
    
    '获取上次复制的位置（存储在隐藏单元格中）
    lastCopiedRow = ws.Range("AA1").Value
    If lastCopiedRow = 0 Then lastCopiedRow = 1 '首次使用
    
    '获取G列数据范围
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    '检查是否有数据
    If lastRow < 2 Then
        MsgBox "没有找到快递单号数据！"
        Exit Sub
    End If
    
    '从上次位置继续收集快递单号
    trackingNumbers = ""
    count = 0
    startRow = 0
    
    For i = lastCopiedRow + 1 To lastRow
        '只处理可见行
        If ws.Rows(i).Hidden = False Then
            Dim cellValue As Variant
            cellValue = ws.Cells(i, "G").Value
            
            '检查单元格是否有值
            If Not IsEmpty(cellValue) And cellValue <> "" Then
                '记录开始行
                If startRow = 0 Then startRow = i
                
                '达到指定数量后停止
                If count >= copyCount Then 
                    '保存当前位置（不包括当前行，下次从这行开始）
                    ws.Range("AA1").Value = i - 1
                    Exit For
                End If
                
                '添加快递单号
                If trackingNumbers <> "" Then
                    trackingNumbers = trackingNumbers & vbCrLf
                End If
                trackingNumbers = trackingNumbers & CStr(cellValue)
                count = count + 1
                
                '更新最后复制的行号
                lastCopiedRow = i
            End If
        End If
    Next i
    
    '如果循环结束还没达到指定数量，说明已经到底了
    If i > lastRow Then
        ws.Range("AA1").Value = lastRow
    End If
    
    '复制到剪贴板
    If trackingNumbers <> "" Then
        '使用数据对象复制到剪贴板
        Dim dataObj As Object
        Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dataObj.SetText trackingNumbers
        dataObj.PutInClipboard
        
        '显示结果信息
        Dim resultMsg As String
        Dim remainingCount As Long
        remainingCount = CountRemainingNumbers(ws, lastCopiedRow, lastRow)
        
        resultMsg = "✅ 复制成功！" & vbCrLf & vbCrLf & _
                   "📋 已复制 " & count & " 个快递单号" & vbCrLf & _
                   "📍 范围：第 " & startRow & " - " & lastCopiedRow & " 行" & vbCrLf & _
                   "📊 剩余可见单号：" & remainingCount & " 个" & vbCrLf & vbCrLf & _
                   "💡 提示：再次运行将从下一个位置继续复制"
        
        MsgBox resultMsg, "操作完成"
    Else
        '没有找到更多单号，询问是否重置
        Dim resetChoice As VbMsgBoxResult
        resetChoice = MsgBox("❌ 没有找到更多快递单号！" & vbCrLf & vbCrLf & _
                            "是否重置复制位置从头开始？", vbYesNo + vbQuestion, "提示")
        
        If resetChoice = vbYes Then
            ws.Range("AA1").Value = 1
            MsgBox "✅ 已重置复制位置，请重新运行！", "重置完成"
        End If
    End If
    
End Sub

Function CountRemainingNumbers(ws As Worksheet, lastCopiedRow As Long, lastRow As Long) As Long
    '计算剩余可见的快递单号数量
    Dim i As Long
    Dim count As Long
    
    count = 0
    For i = lastCopiedRow + 1 To lastRow
        If ws.Rows(i).Hidden = False Then
            If Not IsEmpty(ws.Cells(i, "G").Value) And ws.Cells(i, "G").Value <> "" Then
                count = count + 1
            End If
        End If
    Next i
    
    CountRemainingNumbers = count
End Function

Sub ResetCopyPosition()
    '重置复制位置
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ws.Range("AA1").Value = 1
    MsgBox "✅ 已重置复制位置，下次将从头开始复制！", "重置完成"
End Sub

Sub ShowCopyStatus()
    '显示当前复制状态
    Dim ws As Worksheet
    Dim lastCopiedRow As Long
    Dim remainingCount As Long
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastCopiedRow = ws.Range("AA1").Value
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    If lastCopiedRow = 0 Then lastCopiedRow = 1
    
    remainingCount = CountRemainingNumbers(ws, lastCopiedRow, lastRow)
    
    MsgBox "📊 复制状态信息：" & vbCrLf & vbCrLf & _
           "📍 上次复制到：第 " & lastCopiedRow & " 行" & vbCrLf & _
           "📋 剩余可见单号：" & remainingCount & " 个" & vbCrLf & _
           "📈 数据总行数：" & lastRow & " 行", "复制状态"
End Sub

Sub CopyAsCommaFormat()
    '复制为逗号分隔格式（也支持续复制）
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim trackingNumbers As String
    Dim i As Long
    Dim count As Integer
    Dim copyCount As Integer
    Dim userInput As String
    Dim lastCopiedRow As Long
    
    Set ws = ActiveSheet
    
    userInput = InputBox("请输入要复制的快递单号数量：", "逗号分隔格式", "5")
    
    If userInput = "" Then Exit Sub
    
    If IsNumeric(userInput) Then
        copyCount = CInt(userInput)
    Else
        MsgBox "请输入有效的数字！"
        Exit Sub
    End If
    
    '获取上次复制的位置
    lastCopiedRow = ws.Range("AA1").Value
    If lastCopiedRow = 0 Then lastCopiedRow = 1
    
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    trackingNumbers = ""
    count = 0
    
    For i = lastCopiedRow + 1 To lastRow
        If ws.Rows(i).Hidden = False Then
            Dim cellValue As Variant
            cellValue = ws.Cells(i, "G").Value
            
            If Not IsEmpty(cellValue) And cellValue <> "" Then
                If count >= copyCount Then
                    ws.Range("AA1").Value = i - 1
                    Exit For
                End If
                
                If trackingNumbers <> "" Then
                    trackingNumbers = trackingNumbers & ","
                End If
                trackingNumbers = trackingNumbers & CStr(cellValue)
                count = count + 1
                lastCopiedRow = i
            End If
        End If
    Next i
    
    If i > lastRow Then ws.Range("AA1").Value = lastRow
    
    If trackingNumbers <> "" Then
        Dim dataObj As Object
        Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dataObj.SetText trackingNumbers
        dataObj.PutInClipboard
        
        MsgBox "✅ 已复制 " & count & " 个快递单号（逗号分隔）！", "操作完成"
    Else
        MsgBox "❌ 没有找到更多快递单号！", "提示"
    End If
    
End Sub