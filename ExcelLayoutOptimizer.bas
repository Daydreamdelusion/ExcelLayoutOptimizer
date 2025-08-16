Attribute VB_Name = "ExcelLayoutOptimizer"
'==================================================
' Excel智能布局优化系统 v3.0
' 
' 功能：自动优化Excel表格布局，支持撤销、预览和自定义配置
' 开发：VBA专用，纯单模块解决方案
' 依赖：无外部依赖，仅使用Excel内置功能
' 版本：3.0
' 创建：2024年
'==================================================

Option Explicit

'--------------------------------------------------
' 配置常量区
'--------------------------------------------------
' 列宽边界控制（字符单位）
Private Const DEFAULT_MIN_COLUMN_WIDTH As Double = 8.43     ' 默认最小列宽
Private Const DEFAULT_MAX_COLUMN_WIDTH As Double = 50       ' 默认最大列宽

' 列宽边界控制（像素）
Private Const MIN_COLUMN_WIDTH_PIXELS As Long = 50  ' 最小列宽像素
Private Const MAX_COLUMN_WIDTH_PIXELS As Long = 300 ' 最大列宽像素

' 缓冲区设置（像素）
Private Const TEXT_BUFFER_PIXELS As Long = 15       ' 文本缓冲区
Private Const NUMERIC_BUFFER_PIXELS As Long = 12    ' 数值缓冲区
Private Const DATE_BUFFER_PIXELS As Long = 12       ' 日期缓冲区

' 缓冲区设置（字符单位）
Private Const TEXT_BUFFER_CHARS As Double = 2.0     ' 文本缓冲区
Private Const NUMERIC_BUFFER_CHARS As Double = 1.6  ' 数值缓冲区

' 字符宽度系数
Private Const CHINESE_CHAR_WIDTH_FACTOR As Double = 1.2  ' 中文字符
Private Const ENGLISH_CHAR_WIDTH_FACTOR As Double = 0.6  ' 英文字符
Private Const NUMBER_CHAR_WIDTH_FACTOR As Double = 0.55  ' 数字字符
Private Const OTHER_CHAR_WIDTH_FACTOR As Double = 0.7    ' 其他字符

' 单位转换
Private Const PIXELS_PER_CHAR_UNIT As Double = 7.5  ' 标准字体像素转换

' 行高限制
Private Const MIN_ROW_HEIGHT As Double = 15         ' 最小行高（磅）
Private Const MAX_ROW_HEIGHT As Double = 409        ' 最大行高（磅）

' 性能控制
Private Const MAX_CELLS_LIMIT As Long = 100000      ' 最大处理单元格数
Private Const PROGRESS_UPDATE_INTERVAL As Long = 10 ' 进度更新间隔

' 日期序列号范围
Private Const MIN_EXCEL_DATE As Long = 1            ' Excel最小日期
Private Const MAX_EXCEL_DATE As Long = 2958465      ' Excel最大日期

'--------------------------------------------------
' 数据类型和结构定义
'--------------------------------------------------
' 数据类型枚举
Public Enum DataType
    EmptyCell = 1
    TextValue = 2
    NumericValue = 3
    DateValue = 4
    ErrorValue = 5
    FormulaValue = 6
End Enum

' 字符统计结构
Private Type CharCount
    ChineseCount As Long
    EnglishCount As Long
    NumberCount As Long
    OtherCount As Long
    TotalCount As Long
End Type

' 列宽计算结果
Private Type WidthResult
    FinalWidth As Double      ' 最终列宽（字符单位）
    NeedWrap As Boolean       ' 是否需要换行
    OriginalWidth As Double   ' 原始计算宽度
End Type

' 对齐设置
Private Type AlignmentSettings
    Horizontal As XlHAlign
    Vertical As XlVAlign
End Type

' 列分析结果 - 改为使用数组存储，避免Type限制
Private Type ColumnAnalysisData
    ColumnIndex As Long
    DataType As DataType
    MaxContentWidth As Double
    OptimalWidth As Double
    NeedWrap As Boolean
    CellCount As Long
    HasMergedCells As Boolean
    HasErrors As Boolean
End Type

' 优化统计
Private Type OptimizationStats
    TotalColumns As Long
    AdjustedColumns As Long
    WrapEnabledColumns As Long
    SkippedColumns As Long
    ProcessingTime As Double
    ErrorCount As Long
End Type

'--------------------------------------------------
' Windows API 计时器声明
'--------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

'==================================================
' 公共入口函数
'==================================================

'--------------------------------------------------
' 主入口函数 - 优化选定区域的布局（带配置和预览）
'--------------------------------------------------
Public Sub OptimizeLayout()
    On Error GoTo ErrorHandler
    
    ' 初始化配置
    If Not g_ConfigInitialized Then
        InitializeDefaultConfig
    End If
    
    ' ---- 初始化阶段 ----
    Dim startTime As Long
    startTime = StartTimer()
    
    ' 重置中断标志
    ResetCancelFlag
    
    ' 保存Excel状态
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalStatusBar As Variant
    
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalStatusBar = Application.StatusBar
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' ---- 验证阶段 ----
    Dim selectedRange As Range
    Set selectedRange = Selection
    
    ' 验证选择
    If Not ValidateSelection(selectedRange) Then
        GoTo CleanExit
    End If
    
    ' ---- 配置阶段 ----
    If g_Config.ShowPreview Then
        If Not GetUserConfiguration() Then
            GoTo CleanExit
        End If
    End If
    
    ' ---- 预览阶段 ----
    If g_Config.ShowPreview Then
        Dim previewInfo As PreviewInfo
        previewInfo = CollectPreviewInfo(selectedRange)
        
        Dim userResponse As VbMsgBoxResult
        userResponse = ShowPreviewDialog(previewInfo, selectedRange)
        
        If userResponse <> vbYes Then
            GoTo CleanExit
        End If
    End If
    
    ' ---- 保存撤销信息 ----
    If Not SaveStateForUndo(selectedRange) Then
        MsgBox "无法保存撤销信息，是否继续？", vbYesNo + vbQuestion
        If vbNo Then GoTo CleanExit
    End If
    
    ' ---- 分析阶段 ----
    ShowProgress 0, 100, "正在分析数据..."
    
    ' 读取数据到内存
    Dim dataArray As Variant
    dataArray = ReadRangeToArray(selectedRange)
    
    ' 检查是否包含表头
    Dim hasHeader As Boolean
    If g_Config.SmartHeaderDetection And selectedRange.Rows.Count > 1 Then
        hasHeader = IsHeaderRow(selectedRange.Rows(1), selectedRange.Rows(2))
    End If
    
    ' 分析每列
    Dim columnAnalyses() As ColumnAnalysisData
    columnAnalyses = AnalyzeColumnsWithInterrupt(dataArray, selectedRange)
    
    ' 检查中断
    If g_CancelOperation Then
        GoTo RestoreAndExit
    End If
    
    ' ---- 优化阶段 ----
    ShowProgress 50, 100, "正在应用优化..."
    
    ' 应用列宽优化
    ApplyColumnWidthOptimization selectedRange, columnAnalyses
    
    ' 应用对齐优化（考虑表头）
    ApplyAlignmentOptimizationWithHeader selectedRange, columnAnalyses, hasHeader
    
    ' 应用换行和行高调整
    ApplyWrapAndRowHeight selectedRange, columnAnalyses
    
    ' ---- 完成阶段 ----
    Dim stats As OptimizationStats
    stats = GenerateStatistics(columnAnalyses, GetElapsedTime(startTime))
    
    ShowCompletionMessageWithUndo stats
    
CleanExit:
    ' 恢复Excel状态
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.StatusBar = originalStatusBar
    ClearProgress
    Exit Sub
    
RestoreAndExit:
    ' 用户取消，恢复原始状态
    If g_HasUndoInfo Then
        Application.ScreenUpdating = True
        RestoreFromUndo
    End If
    MsgBox "操作已取消", vbInformation
    GoTo CleanExit
    
ErrorHandler:
    ' 错误处理
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    ClearProgress
    
    MsgBox "优化过程中发生错误：" & vbCrLf & _
           "错误代码：" & Err.Number & vbCrLf & _
           "错误描述：" & Err.Description, _
           vbCritical, "Excel布局优化系统"
    
    Resume CleanExit
End Sub

'--------------------------------------------------
' 快速优化入口（跳过配置和预览）
'--------------------------------------------------
Public Sub QuickOptimize()
    ' 临时禁用预览
    Dim originalShowPreview As Boolean
    originalShowPreview = g_Config.ShowPreview
    g_Config.ShowPreview = False
    
    ' 执行优化
    OptimizeLayout
    
    ' 恢复设置
    g_Config.ShowPreview = originalShowPreview
End Sub

'--------------------------------------------------
' 撤销上次优化操作
'--------------------------------------------------
Public Sub UndoLastOptimization()
    If Not g_HasUndoInfo Then
        MsgBox "没有可撤销的操作", vbInformation, "Excel布局优化系统"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' 执行撤销
    If RestoreFromUndo() Then
        MsgBox "已撤销上次优化操作", vbInformation, "Excel布局优化系统"
        g_HasUndoInfo = False
    Else
        MsgBox "撤销失败", vbCritical, "Excel布局优化系统"
    End If
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "撤销失败：" & Err.Description, vbCritical, "Excel布局优化系统"
End Sub

'==================================================
' 配置管理函数
'==================================================

'--------------------------------------------------
' 初始化默认配置
'--------------------------------------------------
Private Sub InitializeDefaultConfig()
    With g_Config
        .MinColumnWidth = DEFAULT_MIN_COLUMN_WIDTH
        .MaxColumnWidth = DEFAULT_MAX_COLUMN_WIDTH
        .TextBuffer = TEXT_BUFFER_CHARS
        .NumericBuffer = NUMERIC_BUFFER_CHARS
        .WrapThreshold = DEFAULT_MAX_COLUMN_WIDTH
        .SmartHeaderDetection = True
        .ShowPreview = True
        .AutoSave = True
    End With
    g_ConfigInitialized = True
End Sub

'--------------------------------------------------
' 获取用户配置
'--------------------------------------------------
Private Function GetUserConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    Dim response As String
    
    ' 简单配置模式（3个关键参数）
    response = InputBox( _
        "设置最大列宽（字符单位）" & vbCrLf & _
        "范围: 30-100，默认: " & g_Config.MaxColumnWidth & vbCrLf & _
        "直接按Enter使用默认值", _
        "布局优化配置", CStr(g_Config.MaxColumnWidth))
    
    If response = "" Then
        ' 用户按Enter或取消，使用默认值
        GetUserConfiguration = True
        Exit Function
    End If
    
    ' 验证输入
    If IsNumeric(response) Then
        Dim value As Double
        value = CDbl(response)
        If value >= 30 And value <= 100 Then
            g_Config.MaxColumnWidth = value
            g_Config.WrapThreshold = value
        Else
            MsgBox "请输入30-100之间的数值", vbExclamation
            GetUserConfiguration = False
            Exit Function
        End If
    End If
    
    GetUserConfiguration = True
    Exit Function
    
ErrorHandler:
    GetUserConfiguration = False
End Function

'==================================================
' 撤销机制实现
'==================================================

'--------------------------------------------------
' 保存当前状态用于撤销
'--------------------------------------------------
Private Function SaveStateForUndo(targetRange As Range) As Boolean
    On Error GoTo ErrorHandler
    
    ' 初始化撤销信息
    With g_LastUndoInfo
        .RangeAddress = targetRange.Address
        .WorksheetName = targetRange.Worksheet.Name
        .Timestamp = Now
        .Description = "布局优化 " & Format(Now, "hh:mm:ss")
        
        ' 保存列格式
        Dim colCount As Long
        colCount = targetRange.Columns.Count
        ReDim .ColumnFormats(1 To colCount)
        
        Dim i As Long
        For i = 1 To colCount
            With .ColumnFormats(i)
                .ColumnWidth = targetRange.Columns(i).ColumnWidth
                .WrapText = targetRange.Cells(1, i).WrapText
                .HorizontalAlignment = targetRange.Cells(1, i).HorizontalAlignment
                .VerticalAlignment = targetRange.Cells(1, i).VerticalAlignment
            End With
        Next i
        
        ' 保存行高
        Dim rowCount As Long
        rowCount = targetRange.Rows.Count
        ReDim .RowHeights(1 To rowCount)
        
        For i = 1 To rowCount
            .RowHeights(i) = targetRange.Rows(i).RowHeight
        Next i
    End With
    
    g_HasUndoInfo = True
    SaveStateForUndo = True
    Exit Function
    
ErrorHandler:
    SaveStateForUndo = False
End Function

'--------------------------------------------------
' 从撤销信息恢复
'--------------------------------------------------
Private Function RestoreFromUndo() As Boolean
    On Error GoTo ErrorHandler
    
    ' 定位原始区域
    Dim ws As Worksheet
    Set ws = Worksheets(g_LastUndoInfo.WorksheetName)
    Dim targetRange As Range
    Set targetRange = ws.Range(g_LastUndoInfo.RangeAddress)
    
    ' 恢复列格式
    Dim i As Long
    For i = 1 To UBound(g_LastUndoInfo.ColumnFormats)
        With targetRange.Columns(i)
            .ColumnWidth = g_LastUndoInfo.ColumnFormats(i).ColumnWidth
            .WrapText = g_LastUndoInfo.ColumnFormats(i).WrapText
            .HorizontalAlignment = g_LastUndoInfo.ColumnFormats(i).HorizontalAlignment
            .VerticalAlignment = g_LastUndoInfo.ColumnFormats(i).VerticalAlignment
        End With
    Next i
    
    ' 恢复行高
    For i = 1 To UBound(g_LastUndoInfo.RowHeights)
        targetRange.Rows(i).RowHeight = g_LastUndoInfo.RowHeights(i)
    Next i
    
    RestoreFromUndo = True
    Exit Function
    
ErrorHandler:
    RestoreFromUndo = False
End Function

'==================================================
' 预览功能实现
'==================================================

'--------------------------------------------------
' 收集预览信息
'--------------------------------------------------
Private Function CollectPreviewInfo(targetRange As Range) As PreviewInfo
    Dim info As PreviewInfo
    
    With info
        .TotalColumns = targetRange.Columns.Count
        .AffectedCells = targetRange.Cells.Count
        
        ' 快速扫描分析
        Dim col As Range
        Dim maxContent As Double, minContent As Double
        minContent = 999
        maxContent = 0
        
        For Each col In targetRange.Columns
            ' 分析每列内容宽度
            Dim colWidth As Double
            colWidth = AnalyzeColumnWidth(col)
            
            If colWidth < minContent Then minContent = colWidth
            If colWidth > maxContent Then maxContent = colWidth
            
            If Abs(colWidth - col.ColumnWidth) > 0.5 Then
                .ColumnsToAdjust = .ColumnsToAdjust + 1
            End If
            
            If colWidth > g_Config.MaxColumnWidth Then
                .ColumnsNeedWrap = .ColumnsNeedWrap + 1
            End If
        Next col
        
        .MinWidth = minContent
        .MaxWidth = maxContent
        
        ' 检测特殊情况
        .HasMergedCells = HasMergedCells(targetRange)
        .HasFormulas = HasFormulas(targetRange)
        
        ' 估算处理时间（基于经验公式）
        .EstimatedTime = (.AffectedCells / 10000) * 1.5 ' 每万个单元格约1.5秒
        If .EstimatedTime < 0.5 Then .EstimatedTime = 0.5
    End With
    
    CollectPreviewInfo = info
End Function

'--------------------------------------------------
' 分析列宽度（快速版）
'--------------------------------------------------
Private Function AnalyzeColumnWidth(columnRange As Range) As Double
    Dim maxWidth As Double
    Dim cell As Range
    Dim sampleSize As Long
    
    maxWidth = 0
    sampleSize = 0
    
    ' 采样分析（最多100个单元格）
    For Each cell In columnRange.Cells
        If Not IsEmpty(cell.Value) Then
            Dim cellWidth As Double
            cellWidth = CalculateTextWidth(CStr(cell.Value), 11)
            If cellWidth > maxWidth Then maxWidth = cellWidth
            
            sampleSize = sampleSize + 1
            If sampleSize >= 100 Then Exit For
        End If
    Next cell
    
    AnalyzeColumnWidth = maxWidth
End Function

'--------------------------------------------------
' 检查是否包含公式
'--------------------------------------------------
Private Function HasFormulas(checkRange As Range) As Boolean
    On Error Resume Next
    HasFormulas = (checkRange.SpecialCells(xlCellTypeFormulas).Count > 0)
    On Error GoTo 0
End Function

'--------------------------------------------------
' 显示预览对话框
'--------------------------------------------------
Private Function ShowPreviewDialog(info As PreviewInfo, targetRange As Range) As VbMsgBoxResult
    Dim message As String
    
    message = "布局优化预览" & vbCrLf & vbCrLf
    message = message & "优化区域: " & targetRange.Address & vbCrLf
    message = message & String(40, "-") & vbCrLf
    message = message & "• 总列数: " & info.TotalColumns & vbCrLf
    message = message & "• 需调整: " & info.ColumnsToAdjust & " 列" & vbCrLf
    
    If info.ColumnsNeedWrap > 0 Then
        message = message & "• 需换行: " & info.ColumnsNeedWrap & " 列" & vbCrLf
    End If
    
    message = message & "• 宽度范围: " & Format(info.MinWidth, "0.0") & _
              " - " & Format(info.MaxWidth, "0.0") & vbCrLf
    
    If info.HasMergedCells Then
        message = message & "• 警告: 包含合并单元格（将跳过）" & vbCrLf
    End If
    
    If info.HasFormulas Then
        message = message & "• 提示: 包含公式" & vbCrLf
    End If
    
    message = message & String(40, "-") & vbCrLf
    message = message & "预计耗时: " & Format(info.EstimatedTime, "0.0") & " 秒" & vbCrLf & vbCrLf
    message = message & "是否继续？（处理中可按ESC中断）"
    
    ShowPreviewDialog = MsgBox(message, vbYesNoCancel + vbInformation, "Excel布局优化")
End Function

'==================================================
' 智能表头识别
'==================================================

'--------------------------------------------------
' 判断是否为表头行
'--------------------------------------------------
Private Function IsHeaderRow(firstRow As Range, secondRow As Range) As Boolean
    Dim score As Integer
    score = 0
    
    ' 检测标准1：第一行全是文本
    Dim allText As Boolean
    allText = True
    Dim cell As Range
    For Each cell In firstRow.Cells
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
            allText = False
            Exit For
        End If
    Next cell
    If allText Then score = score + 2
    
    ' 检测标准2：第一行无空单元格
    Dim noEmpty As Boolean
    noEmpty = True
    For Each cell In firstRow.Cells
        If IsEmpty(cell.Value) Then
            noEmpty = False
            Exit For
        End If
    Next cell
    If noEmpty Then score = score + 2
    
    ' 检测标准3：格式特征（加粗或背景色）
    Dim hasFormat As Boolean
    For Each cell In firstRow.Cells
        If cell.Font.Bold Or cell.Interior.ColorIndex <> xlNone Then
            hasFormat = True
            Exit For
        End If
    Next cell
    If hasFormat Then score = score + 3
    
    ' 检测标准4：与第二行数据类型差异
    If Not secondRow Is Nothing Then
        Dim typeDiff As Integer
        Dim i As Long
        For i = 1 To Application.Min(firstRow.Cells.Count, secondRow.Cells.Count)
            If GetCellDataType(firstRow.Cells(i).Value) <> _
               GetCellDataType(secondRow.Cells(i).Value) Then
                typeDiff = typeDiff + 1
            End If
        Next i
        If typeDiff > firstRow.Cells.Count / 2 Then score = score + 2
    End If
    
    ' 检测标准5：文本长度
    Dim avgLength As Double
    Dim totalLength As Long
    Dim textCount As Long
    For Each cell In firstRow.Cells
        If Not IsEmpty(cell.Value) Then
            totalLength = totalLength + Len(CStr(cell.Value))
            textCount = textCount + 1
        End If
    Next cell
    If textCount > 0 Then
        avgLength = totalLength / textCount
        If avgLength < 20 Then score = score + 1
    End If
    
    ' 得分>=4认为是表头
    IsHeaderRow = (score >= 4)
End Function

'==================================================
' 中断机制实现
'==================================================

'--------------------------------------------------
' 重置中断标志
'--------------------------------------------------
Private Sub ResetCancelFlag()
    g_CancelOperation = False
    g_CheckCounter = 0
    Application.EnableCancelKey = xlErrorHandler
End Sub

'--------------------------------------------------
' 检查是否需要中断
'--------------------------------------------------
Private Function CheckForCancel() As Boolean
    ' 每100次调用检测一次
    g_CheckCounter = g_CheckCounter + 1
    If g_CheckCounter Mod 100 <> 0 Then
        CheckForCancel = False
        Exit Function
    End If
    
    ' 处理挂起的事件
    DoEvents
    
    ' 检测ESC键
    If g_CancelOperation Then
        If MsgBox("确定要取消当前操作吗？", _
                  vbYesNo + vbQuestion, "取消操作") = vbYes Then
            CheckForCancel = True
        Else
            g_CancelOperation = False
            CheckForCancel = False
        End If
    End If
End Function

'--------------------------------------------------
' 带中断检测的列分析
'--------------------------------------------------
Function AnalyzeColumnsWithInterrupt(dataArray As Variant, targetRange As Range) As ColumnAnalysisData()
    Dim rowCount As Long, colCount As Long
    Dim col As Long
    
    ' 获取数组维度
    If IsArray(dataArray) Then
        rowCount = UBound(dataArray, 1)
        colCount = UBound(dataArray, 2)
    Else
        rowCount = 1
        colCount = 1
    End If
    
    ' 创建列分析数组
    Dim analyses() As ColumnAnalysisData
    ReDim analyses(1 To colCount)
    
    ' 分析每一列
    For col = 1 To colCount
        ShowProgress 20 + (col - 1) * 30 / colCount, 100, "分析列 " & col & "/" & colCount
        
        ' 检查中断
        If CheckForCancel() Then
            g_CancelOperation = True
            AnalyzeColumnsWithInterrupt = analyses
            Exit Function
        End If
        
        analyses(col) = AnalyzeColumn(dataArray, col, rowCount, targetRange.Columns(col))
    Next col
    
    AnalyzeColumnsWithInterrupt = analyses
End Function

'==================================================
' 优化后的格式应用函数
'==================================================

'--------------------------------------------------
' 应用对齐优化（考虑表头）
'--------------------------------------------------
Sub ApplyAlignmentOptimizationWithHeader(targetRange As Range, columnAnalyses() As ColumnAnalysisData, hasHeader As Boolean)
    Dim col As Long
    Dim startRow As Long
    
    If hasHeader Then
        ' 处理表头（第一行）- 统一居中
        For col = 1 To UBound(columnAnalyses)
            If Not columnAnalyses(col).HasMergedCells Then
                With targetRange.Cells(1, col)
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            End If
        Next col
        startRow = 2
    Else
        startRow = 1
    End If
    
    ' 处理数据行
    If targetRange.Rows.Count >= startRow Then
        For col = 1 To UBound(columnAnalyses)
            If Not columnAnalyses(col).HasMergedCells Then
                Dim dataRange As Range
                Set dataRange = targetRange.Cells(startRow, col).Resize(targetRange.Rows.Count - startRow + 1, 1)
                ApplyAlignment dataRange, columnAnalyses(col).DataType, False
            End If
        Next col
    End If
End Sub

'--------------------------------------------------
' 计算最优列宽（使用配置）
'--------------------------------------------------
Function CalculateOptimalWidth(contentWidth As Double, dataType As DataType) As WidthResult
    Dim result As WidthResult
    Dim buffer As Double
    Dim calculatedWidth As Double
    
    ' 根据数据类型确定缓冲区
    Select Case dataType
        Case TextValue
            buffer = g_Config.TextBuffer
        Case NumericValue
            buffer = g_Config.NumericBuffer
        Case DateValue
            buffer = g_Config.NumericBuffer
        Case Else
            buffer = g_Config.TextBuffer
    End Select
    
    ' 计算基础宽度
    calculatedWidth = contentWidth + buffer
    result.OriginalWidth = calculatedWidth
    
    ' 应用边界控制
    If calculatedWidth < g_Config.MinColumnWidth Then
        result.FinalWidth = g_Config.MinColumnWidth
        result.NeedWrap = False
        
    ElseIf calculatedWidth >= g_Config.MaxColumnWidth Then
        result.FinalWidth = g_Config.MaxColumnWidth
        result.NeedWrap = True  ' 标记需要自动换行
        
    Else
        result.FinalWidth = calculatedWidth
        result.NeedWrap = False
    End If
    
    CalculateOptimalWidth = result
End Function

'--------------------------------------------------
' 显示完成消息（带撤销提示）
'--------------------------------------------------
Sub ShowCompletionMessageWithUndo(stats As OptimizationStats)
    Dim message As String
    
    message = "优化完成！" & vbCrLf & vbCrLf & _
              "- 处理列数：" & stats.TotalColumns & " 列" & vbCrLf & _
              "- 调整列数：" & stats.AdjustedColumns & " 列" & vbCrLf & _
              "- 启用换行：" & stats.WrapEnabledColumns & " 列" & vbCrLf & _
              "- 跳过列数：" & stats.SkippedColumns & " 列（含合并单元格）" & vbCrLf & _
              "- 处理时间：" & Format(stats.ProcessingTime, "0.0") & " 秒"
    
    If stats.ErrorCount > 0 Then
        message = message & vbCrLf & "- 错误警告：" & stats.ErrorCount & " 列包含错误值"
    End If
    
    message = message & vbCrLf & vbCrLf & "提示：可使用 UndoLastOptimization 宏撤销本次操作"
    
    MsgBox message, vbInformation, "Excel布局优化系统"
End Sub

'==================================================
' 性能计时器
'==================================================

'--------------------------------------------------
' 开始计时
'--------------------------------------------------
Function StartTimer() As Long
    StartTimer = GetTickCount()
End Function

'--------------------------------------------------
' 获取经过时间
'--------------------------------------------------
Function GetElapsedTime(startTime As Long) As Double
    ' 返回经过的秒数
    GetElapsedTime = (GetTickCount() - startTime) / 1000#
End Function

'==================================================
' 快捷键绑定（更新）
'==================================================

'--------------------------------------------------
' 安装快捷键
'--------------------------------------------------
Sub InstallShortcuts()
    On Error Resume Next
    Application.OnKey "^+{L}", "OptimizeLayout"      ' Ctrl+Shift+L (主功能)
    Application.OnKey "^+{Q}", "QuickOptimize"       ' Ctrl+Shift+Q (快速模式)
    Application.OnKey "^{Z}", "UndoLastOptimization" ' Ctrl+Z (撤销)
    On Error GoTo 0
End Sub

'--------------------------------------------------
' 卸载快捷键
'--------------------------------------------------
Sub UninstallShortcuts()
    On Error Resume Next
    Application.OnKey "^+{L}"
    Application.OnKey "^+{Q}"
    Application.OnKey "^{Z}"
    On Error GoTo 0
End Sub

'==================================================
' 版本信息和帮助
'==================================================

'--------------------------------------------------
' 显示系统信息
'--------------------------------------------------
Sub ShowSystemInfo()
    Dim info As String
    
    info = "Excel智能布局优化系统 v3.0" & vbCrLf & vbCrLf & _
           "新增功能：" & vbCrLf & _
           "✅ 撤销机制 - 支持撤销上次优化" & vbCrLf & _
           "✅ 预览功能 - 优化前显示预览信息" & vbCrLf & _
           "✅ 配置选项 - 可自定义列宽等参数" & vbCrLf & _
           "✅ 智能表头 - 自动识别并特殊处理" & vbCrLf & _
           "✅ 中断机制 - 支持ESC键中断操作" & vbCrLf & vbCrLf & _
           "核心功能：" & vbCrLf & _
           "• 自动计算最优列宽" & vbCrLf & _
           "• 智能数据类型识别" & vbCrLf & _
           "• 标准化对齐方式" & vbCrLf & _
           "• 自动换行和行高调整" & vbCrLf & vbCrLf & _
           "使用方法：" & vbCrLf & _
           "1. 选择需要优化的数据区域" & vbCrLf & _
           "2. 运行 OptimizeLayout 宏" & vbCrLf & _
           "3. 或按快捷键 Ctrl+Shift+L" & vbCrLf & _
           "4. 快速模式 Ctrl+Shift+Q（跳过预览）" & vbCrLf & _
           "5. 撤销操作 运行 UndoLastOptimization" & vbCrLf & vbCrLf & _
           "注意事项：" & vbCrLf & _
           "• 包含合并单元格的区域将被跳过" & vbCrLf & _
           "• 建议在处理前保存文件" & vbCrLf & _
           "• 大数据集处理可能需要较长时间" & vbCrLf & _
           "• 处理过程中可按ESC键中断"
    
    MsgBox info, vbInformation, "关于 Excel布局优化系统"
End Sub
