Sub BatchCopyTrackingNumbers()
    '批量复制快递单号工具（支持续复制）
' --- Configuration ---
Private Const TRACKING_NUMBER_COLUMN As String = "G" '快递单号所在列
Private Const STATE_STORAGE_CELL As String = "AA1"   '用于存储上次复制位置的单元格

' =================================================================
' ==                  PUBLIC FACING MACROS                       ==
' =================================================================

Public Sub BatchCopyTrackingNumbers()
    '批量复制快递单号工具（换行分隔）
    Call ProcessCopy(Separator:=vbCrLf, PromptTitle:="批量复制快递单号")
End Sub

Public Sub CopyAsCommaFormat()
    '复制为逗号分隔格式
    Call ProcessCopy(Separator:=",", PromptTitle:="逗号分隔格式复制")
End Sub

Public Sub ResetCopyPosition()
    '重置复制位置
    On Error Resume Next
    ActiveSheet.Range(STATE_STORAGE_CELL).Value = 1
    On Error GoTo 0
    MsgBox "✅ 已重置复制位置，下次将从头开始复制！", vbInformation, "重置完成"
End Sub

Public Sub ShowCopyStatus()
    '显示当前复制状态
    Dim ws As Worksheet
    Dim lastCopiedRow As Long
    Dim remainingCount As Long
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    On Error Resume Next
    lastCopiedRow = ws.Range(STATE_STORAGE_CELL).Value
    On Error GoTo 0
    
    If lastCopiedRow = 0 Then lastCopiedRow = 1
    
    lastRow = ws.Cells(ws.Rows.Count, TRACKING_NUMBER_COLUMN).End(xlUp).Row
    remainingCount = CountRemainingNumbers(ws, lastCopiedRow, lastRow)
    
    MsgBox "📊 复制状态信息：" & vbCrLf & vbCrLf & _
           "📍 上次复制到：第 " & lastCopiedRow & " 行" & vbCrLf & _
           "📋 剩余可见单号：" & remainingCount & " 个" & vbCrLf & _
           "📈 数据总行数：" & lastRow & " 行", vbInformation, "复制状态"
End Sub

' =================================================================
' ==                  CORE WORKER FUNCTION                       ==
' =================================================================

Private Sub ProcessCopy(Separator As String, PromptTitle As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim trackingNumbers As String
    Dim i As Long
    Dim count As Integer
    Dim copyCount As Integer
    Dim userInput As String
    Dim startRow As Long
    Dim lastCopiedRow As Long
    Dim currentLastCopiedRow As Long
    Dim copyCount As Long '使用 Long 类型更安全
    
    '使用 Collection 提高性能和代码清晰度
    Dim trackingNumbers As Collection
    Set trackingNumbers = New Collection
    
    On Error GoTo ErrorHandler
    Set ws = ActiveSheet
    
    '获取用户输入的复制数量
    userInput = InputBox("请输入要复制的快递单号数量：" & vbCrLf & _
                        "建议范围：1-20个", "批量复制快递单号", "5")
    ' --- 1. 获取用户输入 ---
    Dim userInput As String
    userInput = InputBox("请输入要复制的快递单号数量：", PromptTitle, "5")
    If userInput = "" Then Exit Sub '用户取消
    
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
    copyCount = Val(userInput) 'Val() 比 CInt() 更能容忍无效输入
    If copyCount < 1 Then
        MsgBox "请输入大于0的数字！", vbExclamation
        Exit Sub
    End If
    
    '获取上次复制的位置（存储在隐藏单元格中）
    lastCopiedRow = ws.Range("AA1").Value
    If lastCopiedRow = 0 Then lastCopiedRow = 1 '首次使用
    ' --- 2. 获取状态和数据范围 ---
    On Error Resume Next
    currentLastCopiedRow = ws.Range(STATE_STORAGE_CELL).Value
    On Error GoTo ErrorHandler
    If currentLastCopiedRow = 0 Then currentLastCopiedRow = 1 '首次使用
    
    '获取G列数据范围
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    '检查是否有数据
    lastRow = ws.Cells(ws.Rows.Count, TRACKING_NUMBER_COLUMN).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "没有找到快递单号数据！"
        MsgBox "在 " & TRACKING_NUMBER_COLUMN & " 列没有找到快递单号数据！", vbExclamation
        Exit Sub
    End If
    
    '从上次位置继续收集快递单号
    trackingNumbers = ""
    count = 0
    startRow = 0
    
    For i = lastCopiedRow + 1 To lastRow
        '只处理可见行
        If ws.Rows(i).Hidden = False Then
    ' --- 3. 收集快递单号 ---
    For i = currentLastCopiedRow + 1 To lastRow
        If trackingNumbers.count >= copyCount Then Exit For
        
        If Not ws.Rows(i).Hidden Then
            Dim cellValue As Variant
            cellValue = ws.Cells(i, "G").Value
            cellValue = ws.Cells(i, TRACKING_NUMBER_COLUMN).Value
            
            '检查单元格是否有值
            If Not IsEmpty(cellValue) And cellValue <> "" Then
                '记录开始行
                If startRow = 0 Then startRow = i
                
                '达到指定数量后停止
                If count >= copyCount Then 
                    '保存当前位置（不包括当前行，下次从这行开始）
                    ws.Range("AA1").Value = i - 1
                    Exit For
            ' *** 错误修复：先检查错误，再检查空值 ***
            If Not IsError(cellValue) Then
                If Not IsEmpty(cellValue) And CStr(cellValue) <> "" Then
                    If startRow = 0 Then startRow = i '记录开始行
                    
                    trackingNumbers.Add CStr(cellValue)
                    currentLastCopiedRow = i '更新最后处理的行号
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
    ' --- 4. 更新状态并复制到剪贴板 ---
    ws.Range(STATE_STORAGE_CELL).Value = currentLastCopiedRow
    
    '复制到剪贴板
    If trackingNumbers <> "" Then
        '使用数据对象复制到剪贴板
    If trackingNumbers.count > 0 Then
        '使用 Join 函数，比循环拼接字符串更高效
        Dim tempArray() As String
        ReDim tempArray(0 To trackingNumbers.count - 1)
        Dim j As Long
        For j = 1 To trackingNumbers.count
            tempArray(j - 1) = trackingNumbers(j)
        Next j
        
        Dim resultString As String
        resultString = Join(tempArray, Separator)
        
        '复制到剪贴板
        Dim dataObj As Object
        Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dataObj.SetText trackingNumbers
        dataObj.SetText resultString
        dataObj.PutInClipboard
        
        '显示结果信息
        Dim resultMsg As String
        Dim remainingCount As Long
        remainingCount = CountRemainingNumbers(ws, lastCopiedRow, lastRow)
        remainingCount = CountRemainingNumbers(ws, currentLastCopiedRow, lastRow)
        
        resultMsg = "✅ 复制成功！" & vbCrLf & vbCrLf & _
                   "📋 已复制 " & count & " 个快递单号" & vbCrLf & _
                   "📍 范围：第 " & startRow & " - " & lastCopiedRow & " 行" & vbCrLf & _
                   "📋 已复制 " & trackingNumbers.count & " 个快递单号" & vbCrLf & _
                   "� 范围：第 " & startRow & " - " & currentLastCopiedRow & " 行" & vbCrLf & _
                   "�📊 剩余可见单号：" & remainingCount & " 个" & vbCrLf & vbCrLf & _
                   "💡 提示：再次运行将从下一个位置继续复制"
        
        MsgBox resultMsg, "操作完成"
        MsgBox resultMsg, vbInformation, "操作完成"
    Else
        '没有找到更多单号，询问是否重置
        Dim resetChoice As VbMsgBoxResult
        resetChoice = MsgBox("❌ 没有找到更多快递单号！" & vbCrLf & vbCrLf & _
                            "是否重置复制位置从头开始？", vbYesNo + vbQuestion, "提示")
        
        If resetChoice = vbYes Then
            ws.Range("AA1").Value = 1
            MsgBox "✅ 已重置复制位置，请重新运行！", "重置完成"
            ResetCopyPosition
        End If
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "发生意外错误: " & Err.Description, vbCritical, "错误"
End Sub

Function CountRemainingNumbers(ws As Worksheet, lastCopiedRow As Long, lastRow As Long) As Long
Private Function CountRemainingNumbers(ws As Worksheet, lastCopiedRow As Long, lastRow As Long) As Long
    '计算剩余可见的快递单号数量
    Dim i As Long
    Dim count As Long
    Dim tempCount As Long
    
    count = 0
    tempCount = 0
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
        If Not ws.Rows(i).Hidden Then
            Dim cellValue As Variant
            cellValue = ws.Cells(i, "G").Value
            
            If Not IsEmpty(cellValue) And cellValue <> "" Then
                If count >= copyCount Then
                    ws.Range("AA1").Value = i - 1
                    Exit For
            cellValue = ws.Cells(i, TRACKING_NUMBER_COLUMN).Value
            '同样需要进行错误安全检查
            If Not IsError(cellValue) Then
                If Not IsEmpty(cellValue) And CStr(cellValue) <> "" Then
                    tempCount = tempCount + 1
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
    CountRemainingNumbers = tempCount
End Function