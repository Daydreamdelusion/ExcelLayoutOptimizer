' ================================================================
' 模块：BatchCopyTrackingNumbers_Fixed
' 作用：在筛选后的明细表中，按自上而下可见顺序批量收集快递单号。
' 特性：
'   - 连续扫描数据行，严格遵循界面“可见”顺序（忽略被筛选/手动隐藏的行）。
'   - 通过状态单元格（AA1）跨次记录已复制过的行号，支持“断点续复制”。
'   - 多行输出模式写入专用工作表“复制结果”，避免受原表筛选影响导致可见数量少于期望。
'   - 提供状态查询与一键重置起点功能。
' 使用提示：
'   - 若修改表头行数或快递单号所在列，请同步更新下方常量。
'   - 如需清零进度，运行“ResetCopyPosition”。
' ================================================================
'
' --- 配置常量 ---
Private Const STATE_STORAGE_CELL As String = "AA1"   '用于存储上次复制位置的单元格
Private Const HEADER_ROWS As Long = 1                '表格的表头行数，将从表头下一行开始
' --- 会话变量 ---
Private m_TrackingNumberColumn As String             '缓存本轮会话中用户选择的列，避免重复询问
Private m_RemarksColumn As String                    '缓存本轮会话中用户选择的备注列

' =================================================================
' ==                  PUBLIC FACING MACROS                       ==
' =================================================================

Public Sub BatchCopyTrackingNumbers()
    ' 主入口：批量复制快递单号（以换行分隔）。
    ' - 扫描当前活动工作表中自表头下一行起的可见数据行；
    ' - 跳过空值/错误值、已复制过的行、被筛选/隐藏的行；
    ' - 结果写入“复制结果”工作表，便于一次性全选复制。
    Dim trackingCol As String
    trackingCol = GetTrackingColumn()
    If trackingCol = "" Then Exit Sub ' 用户取消或输入无效
    
    Dim remarksCol As String
    remarksCol = GetRemarksColumn() ' 新增：询问备注列
    
    Call ProcessCopy(Separator:=vbCrLf, PromptTitle:="批量复制快递单号", TrackingColumn:=trackingCol, RemarksColumn:=remarksCol)
End Sub

Public Sub CopyAsCommaFormat()
    ' 复制为单行“逗号分隔”格式（直接写入剪贴板）。
    ' 场景：用于拼接到系统导入框、URL 参数或一行文本中。
    Dim trackingCol As String
    trackingCol = GetTrackingColumn()
    If trackingCol = "" Then Exit Sub ' 用户取消或输入无效
    
    Call ProcessCopy(Separator:=",", PromptTitle:="逗号分隔格式复制", TrackingColumn:=trackingCol, RemarksColumn:="")
End Sub

Public Sub ResetCopyPosition()
    ' 重置复制起点：清除状态单元格的“已复制行号列表”。
    ' 注意：重置后，再次运行将从表头下一行重新开始。
    On Error Resume Next
    ActiveSheet.Range(STATE_STORAGE_CELL).ClearContents
    On Error GoTo 0
    MsgBox "✅ 已重置复制位置，下次将从头开始复制！", vbInformation, "重置完成"
End Sub

Public Sub ShowCopyStatus()
    ' 显示当前复制进度与统计信息：
    ' - 已复制的条数与最后一次复制到的行号；
    ' - 当前筛选条件下剩余可见的可复制数量；
    ' - 数据总行数（用于快速核对数据范围）。
    Dim ws As Worksheet
    Dim copiedRowsDict As Object
    Dim copiedRowsStr As String
    Dim copiedCount As Long
    Dim maxCopiedRow As Long
    Dim remainingCount As Long
    Dim lastRow As Long
    
    Dim trackingCol As String
    trackingCol = GetTrackingColumn()
    If trackingCol = "" Then Exit Sub ' 用户取消或输入无效
    
    Set ws = ActiveSheet
    Set copiedRowsDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    copiedRowsStr = ws.Range(STATE_STORAGE_CELL).Value
    On Error GoTo 0

    If copiedRowsStr <> "" Then
        Dim rowNum As Variant
        For Each rowNum In Split(copiedRowsStr, ",")
            If rowNum <> "" Then
                copiedRowsDict(CStr(rowNum)) = 1
                If CLng(rowNum) > maxCopiedRow Then maxCopiedRow = CLng(rowNum)
            End If
        Next rowNum
    End If
    copiedCount = copiedRowsDict.Count

    lastRow = ws.Cells(ws.Rows.Count, trackingCol).End(xlUp).Row
    remainingCount = CountRemainingNumbers(ws, copiedRowsDict, lastRow, trackingCol)

    Dim statusMsg As String
    If copiedCount = 0 Then
        statusMsg = "尚未开始 (将从第 " & (HEADER_ROWS + 1) & " 行开始)"
    Else
        statusMsg = "已复制 " & copiedCount & " 个 (最后到第 " & maxCopiedRow & " 行)"
    End If

    MsgBox "📊 复制状态信息：" & vbCrLf & vbCrLf & _
           "📍 复制状态：" & statusMsg & vbCrLf & _
           "📋 剩余可见单号：" & remainingCount & " 个" & vbCrLf & _
           "📈 数据总行数：" & lastRow & " 行", vbInformation, "复制状态"
End Sub

' =================================================================
' ==                  CORE WORKER FUNCTION                       ==
' =================================================================

Private Sub ProcessCopy(Separator As String, PromptTitle As String, TrackingColumn As String, RemarksColumn As String)
    ' 核心过程：按“可见行”的自然顺序收集快递单号。
    ' 参数：
    '   - Separator：输出分隔符。vbCrLf 表示多行输出；"," 表示逗号分隔单行。
    '   - PromptTitle：输入框与消息框标题文案。
    '   - TrackingColumn: 包含快递单号的列字母。
    '   - RemarksColumn: (可选)包含备注的列字母。
    ' 逻辑概述：
    '   1) 读取用户期望复制的数量 N；
    '   2) 解析状态单元格（AA1）中已复制行号的集合，用于断点续复制；
    '   3) 自上而下逐行扫描：仅当行“未隐藏且未复制且单元格有效”时计入；
    '   4) 将本批次的新行号追加入状态并持久化；
    '   5) 根据 Separator 输出：
    '        - 多行：写入“复制结果”专用工作表，避免受筛选影响；
    '        - 逗号：直接写入剪贴板；
    '   6) 弹出结果概览并提示下一步操作。
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim trackingNumbers As Collection
    Dim copyCount As Long
    Dim userInput As String
    Dim copiedRowsDict As Object
    Dim copiedRowsStr As String
    
    On Error GoTo ErrorHandler
    Set ws = ActiveSheet
    Set trackingNumbers = New Collection
    Set copiedRowsDict = CreateObject("Scripting.Dictionary")
    
    ' --- 1) 获取用户输入（期望复制数量） ---
    userInput = InputBox("请输入要复制的快递单号数量：", PromptTitle, "10")
    If userInput = "" Then Exit Sub '用户取消
    
    copyCount = Val(userInput)
    If copyCount < 1 Then
        MsgBox "请输入大于0的数字！", vbExclamation
        Exit Sub
    End If
    
    ' --- 调试日志（输出到“立即窗口”）---
    Debug.Print "--------------------------------------------------"
    Debug.Print "Starting New Log at " & Now
    Debug.Print "User requested to copy: " & copyCount & " items."
    ' --- 2) 读取“已复制行号”状态 + 检测数据范围 ---
    On Error Resume Next
    copiedRowsStr = ws.Range(STATE_STORAGE_CELL).Value
    On Error GoTo ErrorHandler

    If copiedRowsStr <> "" Then
        Dim rowNum As Variant
        Debug.Print "Loading previously copied rows from " & STATE_STORAGE_CELL & ": " & copiedRowsStr
        For Each rowNum In Split(copiedRowsStr, ",")
            If rowNum <> "" Then
                copiedRowsDict(CStr(rowNum)) = 1
            End If
        Next rowNum
    End If
    
    Debug.Print "Total previously copied: " & copiedRowsDict.Count & " rows."
    
    lastRow = ws.Cells(ws.Rows.Count, TrackingColumn).End(xlUp).Row
    Debug.Print "Last data row detected in column " & TrackingColumn & ": " & lastRow
    If lastRow <= HEADER_ROWS Then
        MsgBox "在 " & TrackingColumn & " 列没有找到快递单号数据！", vbExclamation
        Debug.Print "No data found. Exiting."
        Exit Sub
    End If
    
    ' --- 3) 收集快递单号（严格按行号升序遍历，保证UI所见顺序） ---
     Dim newlyCopiedRows As New Collection
     Dim i As Long
     Dim cell As Range

     Debug.Print "--- Starting sequential scan from row " & (HEADER_ROWS + 1) & " to " & lastRow & " ---"
     For i = (HEADER_ROWS + 1) To lastRow
         If trackingNumbers.Count >= copyCount Then
             Debug.Print "Reached copy limit of " & copyCount & ". Exiting loop at row " & i
             Exit For
         End If
         
         ' 行可见性判断：通过 EntireRow.Hidden 精准识别筛选/手动隐藏状态。
         Dim isHidden As Boolean
         isHidden = ws.Cells(i, 1).EntireRow.Hidden
         
         If Not isHidden Then
             Debug.Print "Row " & i & " is VISIBLE."
             ' 防重复：跳过已在状态中记录的行号。
             If Not copiedRowsDict.Exists(CStr(i)) Then
                 Debug.Print "  - Row " & i & " has not been copied yet. Processing."
                 Dim cellValue As Variant: cellValue = ws.Cells(i, TrackingColumn).Value
                 Dim cleanValue As String
                 
                 If Not IsError(cellValue) And Not IsEmpty(cellValue) Then
                     ' 通过“SanitizeString”仅保留字母/数字，避免不可见字符（如不间断空格 160）导致剪贴板/拼接异常。
                     cleanValue = SanitizeString(CStr(cellValue))
                     If cleanValue <> "" Then
                         trackingNumbers.Add cleanValue
                         ' 新增逻辑：直接在源工作表的备注列写入默认值
                         If Separator = vbCrLf Then
                             If RemarksColumn <> "" Then
                                 ws.Cells(i, RemarksColumn).Value = "已签收"
                             End If
                         End If
                         
                         newlyCopiedRows.Add CStr(i)
                         Debug.Print "    -> ADDED '" & cleanValue & "'. Total copied now: " & trackingNumbers.Count
                     Else
                         Debug.Print "    -> SKIPPED (cell is empty after cleaning)."
                     End If
                 Else
                     Debug.Print "    -> SKIPPED (cell is empty or contains an error value)."
                 End If
             Else
                 Debug.Print "  - Row " & i & " has already been copied. Skipping."
             End If
         Else
             ' 日志：确认被筛选/隐藏的行已被正确跳过。
             Debug.Print "Row " & i & " is HIDDEN. Skipping."
         End If
     Next i
    
    ' --- 4. 更新状态并复制到剪贴板 ---
    If trackingNumbers.Count > 0 Then
        ' 将新复制的行号追加到状态字符串并保存
        Dim newRow As Variant
        For Each newRow In newlyCopiedRows
            If copiedRowsStr <> "" Then
                copiedRowsStr = copiedRowsStr & "," & newRow
            Else
                copiedRowsStr = newRow
            End If
        Next newRow
        ws.Range(STATE_STORAGE_CELL).Value = copiedRowsStr
        Debug.Print "Updated state in " & STATE_STORAGE_CELL & ": " & copiedRowsStr
        
        ' --- 4) 输出逻辑 ---
        ' 多行模式：写入“复制结果”工作表，避免与源表筛选状态相互影响；
        ' 逗号模式：使用 DataObject 复制到剪贴板，便于粘贴到单行输入场景。
        
        If Separator = vbCrLf Then
            ' 多行输出：写入专用工作表，确保完整可见。
            Dim resultsWs As Worksheet
            Dim createdNew As Boolean
            On Error Resume Next
            Set resultsWs = ThisWorkbook.Worksheets("复制结果")
            
            ' *** 安全检查：防止源工作表与结果工作表重名导致数据被清空 ***
            If Not resultsWs Is Nothing Then
                If resultsWs.Name = ws.Name Then
                    MsgBox "操作中止：源数据工作表不能命名为“复制结果”。" & vbCrLf & vbCrLf & _
                           "请将您的数据工作表重命名，或删除现有的“复制结果”工作表后重试。", vbCritical, "命名冲突"
                    Exit Sub
                End If
            End If
            
            On Error GoTo ErrorHandler
            If resultsWs Is Nothing Then
                Set resultsWs = ThisWorkbook.Worksheets.Add(After:=ws)
                resultsWs.Name = "复制结果"
                createdNew = True
                Debug.Print "Created new results sheet: 复制结果"
            Else
                resultsWs.Cells.Clear
                Debug.Print "Cleared existing results sheet: 复制结果"
            End If
            
            Application.ScreenUpdating = False
            With resultsWs
                .Range("A1").Value = "本次复制的单号 (" & Format(Now, "HH:mm:ss") & ")"
                .Range("A1").Font.Bold = True
                
                Dim j As Long
                For j = 1 To trackingNumbers.Count
                    .Cells(j + 1, 1).Value = trackingNumbers(j)
                Next j
                .Columns("A").AutoFit
            End With
            Application.ScreenUpdating = True
            resultsWs.Activate
            resultsWs.Range("A2:A" & trackingNumbers.Count + 1).Select
            Debug.Print trackingNumbers.Count & " results written to sheet 复制结果"
        Else
            ' 单行输出（逗号分隔）：直接写入剪贴板。
            Dim tempArray() As String: ReDim tempArray(0 To trackingNumbers.Count - 1)
            Dim k As Long: For k = 1 To trackingNumbers.Count: tempArray(k - 1) = trackingNumbers(k): Next k
            Dim resultString As String: resultString = Join(tempArray, Separator)
            
            Dim dataObj As Object: Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
            dataObj.SetText resultString
            dataObj.PutInClipboard
            Debug.Print "Result copied to clipboard via DataObject: " & resultString
        End If
        
        ' --- 5) 结果提示与剩余统计 ---
        Dim resultMsg As String
        Dim remainingCount As Long
        
        ' 将本批次新增的行号回写进字典，便于统计“剩余可见数量”。
        For Each newRow In newlyCopiedRows
            copiedRowsDict(CStr(newRow)) = 1
        Next newRow
        remainingCount = CountRemainingNumbers(ws, copiedRowsDict, lastRow, TrackingColumn)
        
        ' 计算本批次的最小/最大行号用于消息展示。
        Dim firstRowInBatch As Long, lastRowInBatch As Long, tempRow As Long
        firstRowInBatch = CLng(newlyCopiedRows(1))
        lastRowInBatch = CLng(newlyCopiedRows(1))
        For i = 2 To newlyCopiedRows.Count
            tempRow = CLng(newlyCopiedRows(i))
            If tempRow < firstRowInBatch Then firstRowInBatch = tempRow
            If tempRow > lastRowInBatch Then lastRowInBatch = tempRow
        Next i
        
        If Separator = vbCrLf Then
            resultMsg = "✅ 操作完成！" & vbCrLf & vbCrLf & _
                       "📋 已处理 " & trackingNumbers.Count & " 个单号，并在源表备注列标记“已签收”。" & vbCrLf & _
                       "📝 结果已写入工作表『复制结果』" & vbCrLf & _
                       "� 范围：从第 " & firstRowInBatch & " 行到第 " & lastRowInBatch & " 行 (非连续)" & vbCrLf & _
                       "📋 剩余可见单号：" & remainingCount & " 个" & vbCrLf & vbCrLf & _
                       "💡 提示：现在可以从『复制结果』表中复制单号进行查询"
        Else
            resultMsg = "✅ 复制成功！" & vbCrLf & vbCrLf & _
                       "📋 已复制 " & trackingNumbers.Count & " 个快递单号" & vbCrLf & _
                       "📌 范围：从第 " & firstRowInBatch & " 行到第 " & lastRowInBatch & " 行 (非连续)" & vbCrLf & _
                       "📋 剩余可见单号：" & remainingCount & " 个" & vbCrLf & vbCrLf & _
                       "💡 提示：再次运行将从下一个可见位置继续复制"
        End If
        
        MsgBox resultMsg, vbInformation, "操作完成"
    Else
        Debug.Print "No new numbers found to copy."
        '没有找到更多单号，询问是否重置
        Dim resetChoice As VbMsgBoxResult
        resetChoice = MsgBox("❌ 没有找到更多可见的快递单号！" & vbCrLf & vbCrLf & _
                            "是否重置复制位置从头开始？", vbYesNo + vbQuestion, "提示")
        
        If resetChoice = vbYes Then
            Debug.Print "User chose to reset."
            ResetCopyPosition
        End If
    End If
    
    Exit Sub

ErrorHandler:
    ' Ensure screen updating is re-enabled and temp sheet is cleaned up on error
    Application.ScreenUpdating = True
    Debug.Print "!!! An unexpected error occurred: " & Err.Description & " !!!"
    MsgBox "发生意外错误: " & Err.Description, vbCritical, "错误"
End Sub

' =================================================================
' ==                  HELPER FUNCTIONS                           ==
' =================================================================

Private Function CountRemainingNumbers(ws As Worksheet, copiedRowsDict As Object, lastRow As Long, TrackingColumn As String) As Long
    ' 目的：在当前筛选条件下，计算“仍可复制的可见快递单号”数量。
    ' 说明：
    '   - 与主过程一致，逐行判断可见性（EntireRow.Hidden = False），避免 SpecialCells 在复杂筛选场景下的顺序/遗漏问题；
    '   - 跳过已复制过的行（根据状态字典）；
    '   - 仅计入单元格有效（非空、非错误、清洗后非空）的记录。
    Dim tempCount As Long
    Dim i As Long

    Debug.Print "--- Counting remaining numbers ---"
    If lastRow <= HEADER_ROWS Then
        CountRemainingNumbers = 0
        Debug.Print "No data rows to count. Returning 0."
        Exit Function
    End If

    ' 与主过程保持一致的可见性与清洗逻辑，确保统计口径一致。
    tempCount = 0
    For i = (HEADER_ROWS + 1) To lastRow
        If ws.Cells(i, 1).EntireRow.Hidden = False Then ' 行可见性判断
            If Not copiedRowsDict.Exists(CStr(i)) Then
                Dim cellValue As Variant
                cellValue = ws.Cells(i, TrackingColumn).Value
                Dim cleanValue As String
                If Not IsError(cellValue) And Not IsEmpty(cellValue) Then
                    ' 采用相同的清洗策略（仅保留字母/数字）。
                    cleanValue = SanitizeString(CStr(cellValue))
                    If cleanValue <> "" Then
                        tempCount = tempCount + 1
                    End If
                End If
            End If
        End If
    Next i
    
    CountRemainingNumbers = tempCount
    Debug.Print "Counted " & tempCount & " remaining numbers."
    Debug.Print "--- End counting ---"
End Function

Private Function SanitizeString(ByVal inputText As String) As String
    ' 字符清洗：比 WorksheetFunction.Clean + Trim 更稳健。
    ' 规则：逐字符扫描，仅保留英文字母（A-Z、a-z）与数字（0-9），丢弃其他字符。
    ' 目的：避免不可见字符（例如不间断空格 160、控制字符等）造成拼接/剪贴板异常。
    ' 注意：如业务单号可能包含连字符或下划线，请按需扩展保留规则。
    
    Dim outputText As String
    Dim i As Long
    Dim charCode As Integer
    
    outputText = ""
    For i = 1 To Len(inputText)
        charCode = AscW(Mid(inputText, i, 1)) ' 使用 AscW 以兼容 Unicode
        ' 保留：大写(65-90)、小写(97-122)、数字(48-57)
        If (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Or _
           (charCode >= 48 And charCode <= 57) Then
            outputText = outputText & Mid(inputText, i, 1)
        End If
    Next i
    
    SanitizeString = outputText
End Function

Private Function GetTrackingColumn() As String
    ' 获取快递单号列。如果会话中已选择，则直接返回；否则，提示用户输入。
    
    ' 如果本轮会话已指定过列，直接使用，不再询问
    If m_TrackingNumberColumn <> "" Then
        GetTrackingColumn = m_TrackingNumberColumn
        Exit Function
    End If
    
    ' 提示用户输入列字母
    Dim colLetter As String
    colLetter = InputBox("请输入快递单号所在的列字母 (例如: G)", "指定数据列", "G")
    
    If colLetter = "" Then
        GetTrackingColumn = "" ' 用户取消
        Exit Function
    End If
    
    ' 验证输入是否为有效的列地址
    On Error Resume Next
    Dim colNum As Long
    colNum = Range(colLetter & "1").Column
    If Err.Number <> 0 Then
        MsgBox "输入的列字母 '" & colLetter & "' 无效，请确保输入的是单个或多个英文字母。", vbCritical, "输入错误"
        GetTrackingColumn = "" ' 无效输入
    Else
        m_TrackingNumberColumn = UCase(colLetter) ' 缓存选择，并统一为大写
        GetTrackingColumn = m_TrackingNumberColumn
    End If
    On Error GoTo 0
End Function

Private Function GetRemarksColumn() As String
    ' 获取备注列。如果会话中已选择，则直接返回；否则，提示用户输入。
    ' 这是一个可选功能，用户可以留空以禁用。
    
    ' 如果本轮会话已指定过列，直接使用，不再询问
    If m_RemarksColumn <> "" Then
        ' "SKIP" 是一个特殊值，用于记住用户选择不使用此功能
        If m_RemarksColumn = "SKIP" Then
            GetRemarksColumn = ""
        Else
            GetRemarksColumn = m_RemarksColumn
        End If
        Exit Function
    End If
    
    ' 提示用户输入列字母
    Dim colLetter As String
    colLetter = InputBox("（可选）请输入要标记“已签收”的备注列字母 (例如: M)" & vbCrLf & vbCrLf & _
                       "如果留空或取消，将不会标记备注。", "指定备注列", "M")
    
    If colLetter = "" Then
        m_RemarksColumn = "SKIP" ' 缓存用户的“跳过”选择
        GetRemarksColumn = ""    ' 用户跳过
        Exit Function
    End If
    
    ' 验证输入是否为有效的列地址
    On Error Resume Next
    Dim colNum As Long
    colNum = Range(colLetter & "1").Column
    If Err.Number <> 0 Then
        MsgBox "输入的备注列字母 '" & colLetter & "' 无效。请确保输入的是单个或多个英文字母。", vbCritical, "输入错误"
        GetRemarksColumn = "" ' 无效输入
    Else
        m_RemarksColumn = UCase(colLetter) ' 缓存选择，并统一为大写
        GetRemarksColumn = m_RemarksColumn
    End If
    On Error GoTo 0
End Function
