Attribute VB_Name = "ExcelLayoutOptimizer"
'==================================================
' Excel智能布局优化系统 v3.1
' 
' 功能：自动优化Excel表格布局，支持撤销、预览和自定义配置
' 开发：VBA专用，纯单模块解决方案
' 依赖：无外部依赖，仅使用Excel内置VBA功能
' 版本：3.1
' 作者：huangsheng
' 创建：2025年
' 最后更新：2025年8月16日
'
' 修改日志：
' 2025-08-16 v3.1 - huangsheng
'   - 重构代码以彻底解决“行继续标志太多”的编译错误。
'   - 将所有多行 MsgBox 和 InputBox 的字符串拼接重构为变量赋值，提高代码可读性和稳定性。
' 2025-08-16 v3.0 - huangsheng
'   - 添加撤销机制支持
'   - 添加预览功能
'   - 添加配置管理
'   - 添加智能表头识别
'   - 添加中断机制
'   - 修复PreviewInfo类型定义缺失
'   - 修复配置管理全局变量缺失
'   - 修复API声明兼容性问题
' 2025-06-09 v2.0 - huangsheng
'   - 初始版本，核心优化功能
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
Private Const TEXT_BUFFER_CHARS As Double = 2#      ' 文本缓冲区
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

' 列分析结果
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
' v3.0 新增类型定义
'--------------------------------------------------
' 配置参数结构
Private Type OptimizationConfig
    MaxColumnWidth As Double
    MinColumnWidth As Double
    TextBuffer As Double
    NumericBuffer As Double
    WrapThreshold As Double
    SmartHeaderDetection As Boolean
    ShowPreview As Boolean
    AutoSave As Boolean
End Type

' 单元格格式信息
Private Type CellFormat
    ColumnWidth As Double
    WrapText As Boolean
    HorizontalAlignment As XlHAlign
    VerticalAlignment As XlVAlign
    RowHeight As Double
End Type

' 撤销信息
Private Type UndoInfo
    RangeAddress As String
    WorksheetName As String
    ColumnFormats() As CellFormat
    RowHeights() As Double
    Timestamp As Date
    Description As String
End Type

' 预览信息
Private Type PreviewInfo
    TotalColumns As Long
    ColumnsToAdjust As Long
    ColumnsNeedWrap As Long
    MinWidth As Double
    MaxWidth As Double
    EstimatedTime As Double
    AffectedCells As Long
    HasMergedCells As Boolean
    HasFormulas As Boolean
End Type

'--------------------------------------------------
' 全局变量
'--------------------------------------------------
' 配置管理
Private g_Config As OptimizationConfig
Private g_ConfigInitialized As Boolean

' 撤销机制
Private g_LastUndoInfo As UndoInfo
Private g_HasUndoInfo As Boolean

' 中断控制
Private g_CancelOperation As Boolean
Private g_CheckCounter As Long

'--------------------------------------------------
' Windows API 计时器声明
' 注意：修复了条件编译的兼容性问题
'--------------------------------------------------
#If VBA7 Then
    ' 64位Office
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    ' 32位Office
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
        If MsgBox("无法保存撤销信息，是否继续？", vbYesNo + vbQuestion, "Excel布局优化系统") = vbNo Then
            GoTo CleanExit
        End If
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
    MsgBox "操作已取消", vbInformation, "Excel布局优化系统"
    GoTo CleanExit
    
ErrorHandler:
    ' 错误处理
    Dim errorMsg As String
    errorMsg = "优化过程中发生错误：" & vbCrLf
    errorMsg = errorMsg & "错误代码：" & Err.Number & vbCrLf
    errorMsg = errorMsg & "错误描述：" & Err.Description
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    ClearProgress
    
    MsgBox errorMsg, vbCritical, "Excel布局优化系统"
    
    Resume CleanExit
End Sub

'--------------------------------------------------
' 快速优化入口（跳过配置和预览）
'--------------------------------------------------
Public Sub QuickOptimize()
    ' 初始化配置
    If Not g_ConfigInitialized Then
        InitializeDefaultConfig
    End If
    
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
    Dim prompt As String
    
    prompt = "设置最大列宽（字符单位）" & vbCrLf
    prompt = prompt & "范围: 30-100，默认: " & g_Config.MaxColumnWidth & vbCrLf
    prompt = prompt & "直接按Enter使用默认值"
    
    response = InputBox(prompt, "布局优化配置", CStr(g_Config.MaxColumnWidth))
    
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
        If Not IsEmpty(cell.value) Then
            Dim cellWidth As Double
            cellWidth = CalculateTextWidth(CStr(cell.value), 11)
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
    
    message = message & "• 宽度范围: " & Format(info.MinWidth, "0.0") & " - " & Format(info.MaxWidth, "0.0") & vbCrLf
    
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
        If Not IsEmpty(cell.value) And IsNumeric(cell.value) Then
            allText = False
            Exit For
        End If
    Next cell
    If allText Then score = score + 2
    
    ' 检测标准2：第一行无空单元格
    Dim noEmpty As Boolean
    noEmpty = True
    For Each cell In firstRow.Cells
        If IsEmpty(cell.value) Then
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
            If GetCellDataType(firstRow.Cells(i).value) <> GetCellDataType(secondRow.Cells(i).value) Then
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
        If Not IsEmpty(cell.value) Then
            totalLength = totalLength + Len(CStr(cell.value))
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
        If MsgBox("确定要取消当前操作吗？", vbYesNo + vbQuestion, "取消操作") = vbYes Then
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
' 验证和准备函数
'==================================================

'--------------------------------------------------
' 验证用户选择的有效性
'--------------------------------------------------
Function ValidateSelection(selectedRange As Range) As Boolean
    ValidateSelection = False
    
    ' 检查是否有选择
    If selectedRange Is Nothing Then
        MsgBox "请先选择需要优化的区域", vbExclamation, "Excel布局优化系统"
        Exit Function
    End If
    
    ' 检查选择大小
    Dim cellCount As Long
    cellCount = selectedRange.Cells.Count
    
    If cellCount > MAX_CELLS_LIMIT Then
        Dim response As VbMsgBoxResult
        Dim prompt As String
        prompt = "选择区域包含 " & Format(cellCount, "#,##0") & " 个单元格，处理可能需要较长时间。是否继续？"
        response = MsgBox(prompt, vbYesNo + vbQuestion, "Excel布局优化系统")
        If response = vbNo Then Exit Function
    End If
    
    ' 检查工作表保护
    If selectedRange.Worksheet.ProtectContents Then
        MsgBox "工作表受保护，无法进行优化", vbExclamation, "Excel布局优化系统"
        Exit Function
    End If
    
    ' 检查合并单元格
    If HasMergedCells(selectedRange) Then
        Dim mergeResponse As VbMsgBoxResult
        Dim mergePrompt As String
        mergePrompt = "选择区域包含合并单元格，这些区域将被跳过。是否继续？"
        mergeResponse = MsgBox(mergePrompt, vbYesNo + vbQuestion, "Excel布局优化系统")
        If mergeResponse = vbNo Then Exit Function
    End If
    
    ValidateSelection = True
End Function

'--------------------------------------------------
' 批量数据读取（性能优化核心）
'--------------------------------------------------
Function ReadRangeToArray(targetRange As Range) As Variant
    ' 一次性读取，避免循环访问单元格
    Dim dataArray As Variant
    
    ' 使用Value2属性获取原始值（更快）
    dataArray = targetRange.Value2
    
    ' 处理单个单元格的情况
    If Not IsArray(dataArray) Then
        Dim tempArray(1 To 1, 1 To 1) As Variant
        tempArray(1, 1) = dataArray
        dataArray = tempArray
    End If
    
    ReadRangeToArray = dataArray
End Function

'==================================================
' 数据分析函数
'==================================================

'--------------------------------------------------
' 分析单个列的数据特征
'--------------------------------------------------
Function AnalyzeColumn(dataArray As Variant, columnIndex As Long, rowCount As Long, columnRange As Range) As ColumnAnalysisData
    Dim analysis As ColumnAnalysisData
    Dim row As Long
    Dim cellValue As Variant
    Dim cellDataType As DataType
    Dim cellWidth As Double
    Dim maxWidth As Double
    Dim typeCounts(1 To 6) As Long
    
    ' 初始化
    analysis.ColumnIndex = columnIndex
    analysis.CellCount = 0
    analysis.HasMergedCells = HasMergedCells(columnRange)
    analysis.HasErrors = False
    maxWidth = 0
    
    ' 如果包含合并单元格，跳过分析
    If analysis.HasMergedCells Then
        analysis.DataType = TextValue
        analysis.MaxContentWidth = 0
        analysis.OptimalWidth = 0
        analysis.NeedWrap = False
        AnalyzeColumn = analysis
        Exit Function
    End If
    
    ' 分析每个单元格
    For row = 1 To rowCount
        If IsArray(dataArray) And UBound(dataArray, 2) >= columnIndex Then
            cellValue = dataArray(row, columnIndex)
        Else
            cellValue = dataArray
        End If
        
        If Not IsEmpty(cellValue) And cellValue <> "" Then
            analysis.CellCount = analysis.CellCount + 1
            
            ' 获取数据类型
            cellDataType = GetCellDataType(cellValue)
            typeCounts(cellDataType) = typeCounts(cellDataType) + 1
            
            ' 检查错误值
            If cellDataType = ErrorValue Then
                analysis.HasErrors = True
            End If
            
            ' 计算单元格宽度
            If cellDataType <> ErrorValue Then
                cellWidth = CalculateTextWidth(SafeGetCellValue(cellValue), 11)
                If cellWidth > maxWidth Then
                    maxWidth = cellWidth
                End If
            End If
        End If
    Next row
    
    ' 确定列的主导数据类型
    analysis.DataType = DetermineColumnType(typeCounts)
    
    ' 计算最优列宽
    analysis.MaxContentWidth = maxWidth
    Dim widthResult As WidthResult
    widthResult = CalculateOptimalWidth(maxWidth, analysis.DataType)
    analysis.OptimalWidth = widthResult.FinalWidth
    analysis.NeedWrap = widthResult.NeedWrap
    
    AnalyzeColumn = analysis
End Function

'--------------------------------------------------
' 判断单个单元格的数据类型
'--------------------------------------------------
Function GetCellDataType(cellValue As Variant) As DataType
    ' 优先级顺序：错误值 > 空值 > 日期 > 数值 > 文本
    
    ' 1. 错误值检测（最高优先级）
    If IsError(cellValue) Then
        GetCellDataType = ErrorValue
        Exit Function
    End If
    
    ' 2. 空值检测
    If IsEmpty(cellValue) Or cellValue = "" Then
        GetCellDataType = EmptyCell
        Exit Function
    End If
    
    ' 3. 日期检测（必须在数值检测之前）
    If IsDate(cellValue) Then
        ' 额外验证：Excel日期序列号范围 1-2958465
        If IsNumeric(cellValue) Then
            Dim numValue As Double
            numValue = CDbl(cellValue)
            If numValue >= MIN_EXCEL_DATE And numValue <= MAX_EXCEL_DATE Then
                GetCellDataType = DateValue
                Exit Function
            End If
        End If
        GetCellDataType = DateValue
        Exit Function
    End If
    
    ' 4. 数值检测
    If IsNumeric(cellValue) Then
        GetCellDataType = NumericValue
        Exit Function
    End If
    
    ' 5. 默认为文本
    GetCellDataType = TextValue
End Function

'--------------------------------------------------
' 确定列的主导数据类型
'--------------------------------------------------
Function DetermineColumnType(typeCounts() As Long) As DataType
    Dim maxCount As Long
    Dim dominantType As DataType
    Dim i As Long
    
    ' 找出主导类型（忽略空值和错误值）
    maxCount = 0
    dominantType = TextValue  ' 默认文本
    
    For i = 1 To 6
        ' 跳过空值和错误值
        If i <> EmptyCell And i <> ErrorValue Then
            If typeCounts(i) > maxCount Then
                maxCount = typeCounts(i)
                dominantType = i
            End If
        End If
    Next i
    
    ' 特殊规则：如果有任何文本，整列按文本处理
    If typeCounts(TextValue) > 0 Then
        dominantType = TextValue
    End If
    
    DetermineColumnType = dominantType
End Function

'==================================================
' 宽度计算函数
'==================================================

'--------------------------------------------------
' 计算文本显示宽度
'--------------------------------------------------
Function CalculateTextWidth(text As String, fontSize As Single) As Double
    Dim charCounts As CharCount
    Dim pixelWidth As Double
    Dim chineseWidth As Double
    Dim englishWidth As Double
    Dim numberWidth As Double
    Dim otherWidth As Double
    
    ' 统计各类字符
    charCounts = CountCharTypes(text)
    
    ' 分别计算各类字符的像素宽度
    chineseWidth = charCounts.ChineseCount * CHINESE_CHAR_WIDTH_FACTOR
    englishWidth = charCounts.EnglishCount * ENGLISH_CHAR_WIDTH_FACTOR
    numberWidth = charCounts.NumberCount * NUMBER_CHAR_WIDTH_FACTOR
    otherWidth = charCounts.OtherCount * OTHER_CHAR_WIDTH_FACTOR
    
    ' 计算总像素宽度
    pixelWidth = (chineseWidth + englishWidth + numberWidth + otherWidth) * fontSize
    
    ' 转换为字符单位
    CalculateTextWidth = pixelWidth / PIXELS_PER_CHAR_UNIT
End Function

'--------------------------------------------------
' 统计字符类型
'--------------------------------------------------
Function CountCharTypes(text As String) As CharCount
    Dim result As CharCount
    Dim i As Long
    Dim charCode As Long
    
    For i = 1 To Len(text)
        charCode = AscW(Mid(text, i, 1))
        
        If charCode >= &H4E00 And charCode <= &H9FFF Then
            ' 中文字符 (CJK统一汉字)
            result.ChineseCount = result.ChineseCount + 1
        ElseIf charCode >= 48 And charCode <= 57 Then
            ' 数字 (0-9)
            result.NumberCount = result.NumberCount + 1
        ElseIf (charCode >= 65 And charCode <= 90) Or (charCode >= 97 And charCode <= 122) Then
            ' 英文字母 (A-Z, a-z)
            result.EnglishCount = result.EnglishCount + 1
        Else
            ' 其他字符（标点、符号等）
            result.OtherCount = result.OtherCount + 1
        End If
    Next i
    
    result.TotalCount = result.ChineseCount + result.EnglishCount + result.NumberCount + result.OtherCount
    CountCharTypes = result
End Function

'--------------------------------------------------
' 计算最优列宽（使用配置）
'--------------------------------------------------
Function CalculateOptimalWidth(contentWidth As Double, dataType As DataType) As WidthResult
    Dim result As WidthResult
    Dim buffer As Double
    Dim calculatedWidth As Double
    
    ' 确保配置已初始化
    If Not g_ConfigInitialized Then
        InitializeDefaultConfig
    End If
    
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

'==================================================
' 格式应用函数
'==================================================

'--------------------------------------------------
' 应用列宽优化
'--------------------------------------------------
Sub ApplyColumnWidthOptimization(targetRange As Range, columnAnalyses() As ColumnAnalysisData)
    Dim col As Long
    
    For col = 1 To UBound(columnAnalyses)
        ' 跳过包含合并单元格的列
        If Not columnAnalyses(col).HasMergedCells And columnAnalyses(col).OptimalWidth > 0 Then
            targetRange.Columns(col).ColumnWidth = columnAnalyses(col).OptimalWidth
        End If
    Next col
End Sub

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
' 应用对齐设置
'--------------------------------------------------
Sub ApplyAlignment(targetRange As Range, dataType As DataType, isHeader As Boolean)
    Dim settings As AlignmentSettings
    settings = GetAlignmentForDataType(dataType, isHeader)
    
    ' 批量应用，提高性能
    With targetRange
        If settings.Horizontal <> xlGeneral Then
            .HorizontalAlignment = settings.Horizontal
        End If
        If settings.Vertical <> xlBottom Then
            .VerticalAlignment = settings.Vertical
        End If
    End With
End Sub

'--------------------------------------------------
' 获取数据类型对应的对齐方式
'--------------------------------------------------
Function GetAlignmentForDataType(dataType As DataType, isHeader As Boolean) As AlignmentSettings
    Dim settings As AlignmentSettings
    
    If isHeader Then
        ' 表头统一居中
        settings.Horizontal = xlCenter
        settings.Vertical = xlCenter
    Else
        Select Case dataType
            Case TextValue
                settings.Horizontal = xlLeft
                settings.Vertical = xlCenter
                
            Case NumericValue, DateValue
                settings.Horizontal = xlRight
                settings.Vertical = xlCenter
                
            Case EmptyCell
                ' 保持默认，不修改
                settings.Horizontal = xlGeneral
                settings.Vertical = xlBottom
                
            Case Else
                ' 其他情况保持默认
                settings.Horizontal = xlGeneral
                settings.Vertical = xlCenter
        End Select
    End If
    
    GetAlignmentForDataType = settings
End Function

'--------------------------------------------------
' 应用换行和行高调整
'--------------------------------------------------
Sub ApplyWrapAndRowHeight(targetRange As Range, columnAnalyses() As ColumnAnalysisData)
    Dim col As Long
    Dim needRowHeightAdjust As Boolean
    
    needRowHeightAdjust = False
    
    ' 设置需要换行的列
    For col = 1 To UBound(columnAnalyses)
        If Not columnAnalyses(col).HasMergedCells And columnAnalyses(col).NeedWrap Then
            targetRange.Columns(col).WrapText = True
            needRowHeightAdjust = True
        End If
    Next col
    
    ' 如果有列启用了换行，调整行高
    If needRowHeightAdjust Then
        AdjustRowHeightForWrap targetRange
    End If
End Sub

'--------------------------------------------------
' 调整行高以适应换行
'--------------------------------------------------
Sub AdjustRowHeightForWrap(targetRange As Range)
    Dim row As Range
    Application.ScreenUpdating = False
    
    ' 逐行调整，避免影响未选择区域
    For Each row In targetRange.Rows
        row.EntireRow.AutoFit
        
        ' 设置最小和最大行高限制
        If row.RowHeight < MIN_ROW_HEIGHT Then
            row.RowHeight = MIN_ROW_HEIGHT
        ElseIf row.RowHeight > MAX_ROW_HEIGHT Then
            row.RowHeight = MAX_ROW_HEIGHT
        End If
    Next row
    
    Application.ScreenUpdating = True
End Sub

'==================================================
' 辅助工具函数
'==================================================

'--------------------------------------------------
' 安全获取单元格值
'--------------------------------------------------
Function SafeGetCellValue(cellValue As Variant) As String
    On Error Resume Next
    
    If IsError(cellValue) Then
        SafeGetCellValue = ""  ' 错误值返回空字符串
    ElseIf IsNull(cellValue) Then
        SafeGetCellValue = ""
    ElseIf IsEmpty(cellValue) Then
        SafeGetCellValue = ""
    Else
        SafeGetCellValue = CStr(cellValue)
    End If
    
    On Error GoTo 0
End Function

'--------------------------------------------------
' 检测是否包含合并单元格
'--------------------------------------------------
Function HasMergedCells(checkRange As Range) As Boolean
    Dim cell As Range
    
    On Error Resume Next
    For Each cell In checkRange
        If cell.MergeCells Then
            HasMergedCells = True
            Exit Function
        End If
    Next cell
    On Error GoTo 0
    
    HasMergedCells = False
End Function

'--------------------------------------------------
' 生成优化统计信息
'--------------------------------------------------
Function GenerateStatistics(columnAnalyses() As ColumnAnalysisData, processingTime As Double) As OptimizationStats
    Dim stats As OptimizationStats
    Dim col As Long
    
    stats.TotalColumns = UBound(columnAnalyses)
    stats.ProcessingTime = processingTime
    
    For col = 1 To UBound(columnAnalyses)
        If columnAnalyses(col).HasMergedCells Then
            stats.SkippedColumns = stats.SkippedColumns + 1
        Else
            stats.AdjustedColumns = stats.AdjustedColumns + 1
            
            If columnAnalyses(col).NeedWrap Then
                stats.WrapEnabledColumns = stats.WrapEnabledColumns + 1
            End If
        End If
        
        If columnAnalyses(col).HasErrors Then
            stats.ErrorCount = stats.ErrorCount + 1
        End If
    Next col
    
    GenerateStatistics = stats
End Function

'==================================================
' 进度显示和用户反馈
'==================================================

'--------------------------------------------------
' 显示进度
'--------------------------------------------------
Sub ShowProgress(current As Long, total As Long, message As String)
    Dim percentage As Long
    Dim progressBar As String
    Dim i As Long
    
    percentage = CLng((current / total) * 100)
    
    ' 创建文本进度条
    progressBar = "["
    For i = 1 To 20
        If i <= (percentage / 5) Then
            progressBar = progressBar & "="
        Else
            progressBar = progressBar & " "
        End If
    Next i
    progressBar = progressBar & "]"
    
    Application.StatusBar = message & " " & progressBar & " " & percentage & "%"
    
    ' 定期刷新显示
    If current Mod PROGRESS_UPDATE_INTERVAL = 0 Then
        DoEvents
    End If
End Sub

'--------------------------------------------------
' 清除进度显示
'--------------------------------------------------
Sub ClearProgress()
    Application.StatusBar = False
End Sub

'--------------------------------------------------
' 显示完成消息（带撤销提示）
'--------------------------------------------------
Sub ShowCompletionMessageWithUndo(stats As OptimizationStats)
    Dim message As String
    
    message = "优化完成！" & vbCrLf & vbCrLf
    message = message & "- 处理列数：" & stats.TotalColumns & " 列" & vbCrLf
    message = message & "- 调整列数：" & stats.AdjustedColumns & " 列" & vbCrLf
    message = message & "- 启用换行：" & stats.WrapEnabledColumns & " 列" & vbCrLf
    message = message & "- 跳过列数：" & stats.SkippedColumns & " 列（含合并单元格）" & vbCrLf
    message = message & "- 处理时间：" & Format(stats.ProcessingTime, "0.0") & " 秒"
    
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
    
    info = "Excel智能布局优化系统 v3.1" & vbCrLf & vbCrLf
    info = info & "新增功能：" & vbCrLf
    info = info & "✅ 撤销机制 - 支持撤销上次优化" & vbCrLf
    info = info & "✅ 预览功能 - 优化前显示预览信息" & vbCrLf
    info = info & "✅ 配置选项 - 可自定义列宽等参数" & vbCrLf
    info = info & "✅ 智能表头 - 自动识别并特殊处理" & vbCrLf
    info = info & "✅ 中断机制 - 支持ESC键中断操作" & vbCrLf & vbCrLf
    info = info & "核心功能：" & vbCrLf
    info = info & "• 自动计算最优列宽" & vbCrLf
    info = info & "• 智能数据类型识别" & vbCrLf
    info = info & "• 标准化对齐方式" & vbCrLf
    info = info & "• 自动换行和行高调整" & vbCrLf & vbCrLf
    info = info & "使用方法：" & vbCrLf
    info = info & "1. 选择需要优化的数据区域" & vbCrLf
    info = info & "2. 运行 OptimizeLayout 宏" & vbCrLf
    info = info & "3. 或按快捷键 Ctrl+Shift+L" & vbCrLf
    info = info & "4. 快速模式 Ctrl+Shift+Q（跳过预览）" & vbCrLf
    info = info & "5. 撤销操作 运行 UndoLastOptimization" & vbCrLf & vbCrLf
    info = info & "注意事项：" & vbCrLf
    info = info & "• 包含合并单元格的区域将被跳过" & vbCrLf
    info = info & "• 建议在处理前保存文件" & vbCrLf
    info = info & "• 大数据集处理可能需要较长时间" & vbCrLf
    info = info & "• 处理过程中可按ESC键中断" & vbCrLf & vbCrLf
    info = info & "作者：hsd" & vbCrLf
    info = info & "版本：3.1" & vbCrLf
    info = info & "更新日期：2025年8月16日"
    
    MsgBox info, vbInformation, "关于 Excel布局优化系统"
End Sub
