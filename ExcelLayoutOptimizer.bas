'==================================================
' Excel智能布局优化系统 v3.2
'
' 功能：自动优化Excel表格布局，支持撤销、预览和自定义配置
' 特性：标题优先显示、分块处理、缓存优化、智能识别、错误分级、隐藏行列保护
' 依赖：无外部依赖，仅使用Excel内置VBA功能
' 作者：dadada
' 创建：2025年
' 最后更新：2025年8月18日
'
' 修改日志：
' 2025-08-18 v3.2 - dadada
'   - 新增隐藏行列保护机制
'   - 优化时不会取消用户的隐藏设置
'   - 仅优化可见范围内的布局
'   - 新增隐藏行列保护测试用例
' 2025-08-18 v3.1 - dadada
'   - 新增标题优先完整显示功能
'   - 标题自动换行和行高调整
'   - 完善预览信息显示标题相关内容
'   - 增加标题优先功能测试用例
' 2025-08-16 v3.0 - dadada
'   - 增加分块处理和缓存机制
'   - 完善数据类型智能识别
'   - 实现分级错误处理
'   - 增加配置持久化
'   - 添加测试套件
' 2025-08-16 v2.1 - dadada
'   - 重构代码解决编译错误
'   - 添加撤销、预览、配置功能
' 2025-08-09 v1.0 - dadada
'   - 初始版本
'==================================================

Option Explicit

'--------------------------------------------------
' 配置常量区
'--------------------------------------------------
' 列宽边界控制（字符单位）
Private Const DEFAULT_MIN_COLUMN_WIDTH As Double = 8.43
Private Const DEFAULT_MAX_COLUMN_WIDTH As Double = 70  ' 增加到70以支持更长的标题

' 超长文本处理常量（新增）
Private Const EXTREME_TEXT_WIDTH As Double = 120        ' 极长文本固定宽度
Private Const LONG_TEXT_THRESHOLD As Long = 100         ' 长文本阈值（字符数）
Private Const VERY_LONG_TEXT_THRESHOLD As Long = 200    ' 极长文本阈值（字符数）
Private Const MAX_WRAP_LINES As Long = 3                ' 最大换行行数（限制为3行避免过高）

' 列宽边界控制（像素）
Private Const MIN_COLUMN_WIDTH_PIXELS As Long = 50
Private Const MAX_COLUMN_WIDTH_PIXELS As Long = 300

' 缓冲区设置（像素）
Private Const TEXT_BUFFER_PIXELS As Long = 15
Private Const NUMERIC_BUFFER_PIXELS As Long = 12
Private Const DATE_BUFFER_PIXELS As Long = 12

' 缓冲区设置（字符单位）
Private Const TEXT_BUFFER_CHARS As Double = 2.0    ' 从3.5减少到2.0，避免过度缓冲
Private Const NUMERIC_BUFFER_CHARS As Double = 1.6
Private Const DATE_BUFFER_CHARS As Double = 2#      ' 添加日期缓冲区设置

' 字符宽度系数
Private Const CHINESE_CHAR_WIDTH_FACTOR As Double = 1.2
Private Const ENGLISH_CHAR_WIDTH_FACTOR As Double = 0.6
Private Const NUMBER_CHAR_WIDTH_FACTOR As Double = 0.55
Private Const OTHER_CHAR_WIDTH_FACTOR As Double = 0.7

' 单位转换
Private Const PIXELS_PER_CHAR_UNIT As Double = 7.5

' 行高限制
Private Const MIN_ROW_HEIGHT As Double = 15
Private Const MAX_ROW_HEIGHT As Double = 409

' 性能控制
Private Const MAX_CELLS_LIMIT As Long = 100000
Private Const PROGRESS_UPDATE_INTERVAL As Long = 10
Private Const CHUNK_SIZE As Long = 1000
Private Const CACHE_SIZE As Long = 100

' 日期序列号范围
Private Const MIN_EXCEL_DATE As Long = 1
Private Const MAX_EXCEL_DATE As Long = 2958465

'--------------------------------------------------
' 数据类型和结构定义
'--------------------------------------------------
' 数据类型枚举（细化版）
Public Enum DataType
    EmptyCell = 1
    ShortText = 2
    MediumText = 3
    LongText = 4
    IntegerValue = 5
    DecimalValue = 6
    CurrencyValue = 7
    PercentageValue = 8
    DateValue = 9
    TimeValue = 10
    DateTimeValue = 11
    BooleanValue = 12
    ErrorValue = 13
    FormulaValue = 14
    MixedContent = 15
End Enum

' 文本长度分级枚举（新增）
Public Enum TextLengthCategory
    ShortText = 1      ' <= 20字符
    MediumText = 2     ' 21-50字符
    LongText = 3       ' 51-100字符
    VeryLongText = 4   ' 101-200字符
    ExtremeText = 5    ' > 200字符
End Enum

' 错误级别枚举
Public Enum ErrorLevel
    Fatal = 1
    Severe = 2
    Warning = 3
    Info = 4
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
    FinalWidth As Double
    NeedWrap As Boolean
    OriginalWidth As Double
End Type

' 智能换行布局结果
Private Type WrapLayout
    TotalLines As Long
    OptimalRowHeight As Double
    BreakPoints() As Long
    NeedWrap As Boolean
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
    TypeDistribution(1 To 15) As Long
    ' 标题相关新增字段
    HeaderText As String
    HeaderWidth As Double
    HeaderNeedWrap As Boolean
    HeaderRowHeight As Double
    IsHeaderColumn As Boolean
End Type

' 优化统计
Private Type OptimizationStats
    TotalColumns As Long
    AdjustedColumns As Long
    WrapEnabledColumns As Long
    SkippedColumns As Long
    ProcessingTime As Double
    ErrorCount As Long
    CacheHits As Long
    ChunksProcessed As Long
End Type

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
    UseCache As Boolean
    ChunkProcessing As Boolean
    ' 标题相关新增配置
    HeaderPriority As Boolean  ' 标题优先模式
    HeaderMaxWrapLines As Long ' 标题最大换行数
    HeaderMinHeight As Double  ' 标题最小行高
    ' 超长文本处理配置（新增）
    ExtremeTextWidth As Double     ' 极长文本固定宽度
    LongTextThreshold As Long      ' 长文本阈值（字符数）
    SmartLineBreak As Boolean      ' 智能断行开关
    MaxWrapLines As Long          ' 最大换行行数
    LongTextExtendThreshold As Long ' 长文本扩展阈值
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

' 缓存结构
Private Type CellWidthCache
    Content As String
    Width As Double
    Hits As Long
End Type

' 错误信息结构
Private Type ErrorInfo
    Level As ErrorLevel
    Code As Long
    Description As String
    Action As String
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

' 缓存管理
Private g_WidthCache() As CellWidthCache
Private g_CacheSize As Long
Private g_CacheHits As Long

' 性能统计
Private g_ChunksProcessed As Long

'--------------------------------------------------
' Windows API 声明
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
' 主入口函数 - 优化选定区域的布局
'--------------------------------------------------
Public Sub OptimizeLayout()
    On Error GoTo ErrorHandler
    
    ' 初始化
    If Not g_ConfigInitialized Then
        InitializeDefaultConfig
        LoadConfigFromWorkbook
    End If
    
    InitializeCache
    ResetCancelFlag
    g_ChunksProcessed = 0
    
    Dim startTime As Long
    startTime = StartTimer()
    
    ' 保存Excel状态
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalStatusBar As Variant
    
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalStatusBar = Application.StatusBar
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 验证选择
    Dim selectedRange As Range
    Set selectedRange = Selection
    
    If Not ValidateSelectionEnhanced(selectedRange) Then
        GoTo CleanExit
    End If
    
    ' 配置阶段
    If g_Config.ShowPreview Then
        If Not GetUserConfiguration() Then
            GoTo CleanExit
        End If
    End If
    
    ' 预览阶段
    If g_Config.ShowPreview Then
        Dim previewInfo As PreviewInfo
        previewInfo = CollectPreviewInfo(selectedRange)
        
        If ShowPreviewDialog(previewInfo, selectedRange) <> vbYes Then
            GoTo CleanExit
        End If
    End If
    
    ' 保存撤销信息
    If Not SaveStateForUndo(selectedRange) Then
        If MsgBox("无法保存撤销信息，是否继续？", vbYesNo + vbQuestion, "Excel布局优化系统") = vbNo Then
            GoTo CleanExit
        End If
    End If
    
    ' 执行优化
    Dim success As Boolean
    If g_Config.ChunkProcessing And selectedRange.Rows.Count > CHUNK_SIZE Then
        success = ProcessInChunks(selectedRange)
    Else
        success = ProcessNormal(selectedRange)
    End If
    
    If Not success Then
        GoTo RestoreAndExit
    End If
    
    ' 保存配置
    If g_Config.AutoSave Then
        SaveConfigToWorkbook
    End If
    
    ' 显示统计
    Dim stats As OptimizationStats
    stats = GenerateEnhancedStatistics(selectedRange, GetElapsedTime(startTime))
    ShowCompletionMessageEnhanced stats
    
CleanExit:
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.StatusBar = originalStatusBar
    ClearProgress
    ClearCache
    Exit Sub
    
RestoreAndExit:
    If g_HasUndoInfo Then
        Application.ScreenUpdating = True
        RestoreFromUndo
    End If
    MsgBox "操作已取消或失败", vbInformation, "Excel布局优化系统"
    GoTo CleanExit
    
ErrorHandler:
    Dim errorInfo As ErrorInfo
    errorInfo = ClassifyError(Err.Number, Err.Description)
    HandleErrorByLevel errorInfo
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    ClearProgress
    ClearCache
    
    Resume CleanExit
End Sub

'--------------------------------------------------
' 快速优化入口
'--------------------------------------------------
Public Sub QuickOptimize()
    If Not g_ConfigInitialized Then
        InitializeDefaultConfig
    End If
    
    Dim originalShowPreview As Boolean
    originalShowPreview = g_Config.ShowPreview
    g_Config.ShowPreview = False
    
    OptimizeLayout
    
    g_Config.ShowPreview = originalShowPreview
End Sub

'--------------------------------------------------
' 保守优化入口（新增）- 避免行高过度调整
'--------------------------------------------------
Public Sub ConservativeOptimize()
    If Not g_ConfigInitialized Then
        InitializeDefaultConfig
    End If
    
    ' 保存原始配置
    Dim originalConfig As OptimizationConfig
    originalConfig = g_Config
    
    ' 使用保守设置
    g_Config.ShowPreview = False
    g_Config.HeaderPriority = False       ' 关闭标题优先，避免过度行高调整
    g_Config.SmartLineBreak = False       ' 关闭智能断行
    g_Config.MaxWrapLines = 2             ' 最多2行换行
    g_Config.HeaderMaxWrapLines = 2       ' 标题最多2行换行
    
    ' 执行优化
    OptimizeLayout
    
    ' 恢复原始配置
    g_Config = originalConfig
End Sub

'--------------------------------------------------
' 撤销上次优化
'--------------------------------------------------
Public Sub UndoLastOptimization()
    If Not g_HasUndoInfo Then
        MsgBox "没有可撤销的操作", vbInformation, "Excel布局优化系统"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
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
' 分块处理实现
'==================================================

'--------------------------------------------------
' 分块处理主函数
'--------------------------------------------------
Private Function ProcessInChunks(targetRange As Range) As Boolean
    On Error GoTo ErrorHandler
    
    Dim totalRows As Long
    totalRows = targetRange.Rows.Count
    
    Dim startRow As Long, endRow As Long
    Dim chunkIndex As Long
    chunkIndex = 0
    
    For startRow = 1 To totalRows Step CHUNK_SIZE
        endRow = Application.Min(startRow + CHUNK_SIZE - 1, totalRows)
        chunkIndex = chunkIndex + 1
        
        ShowProgress startRow, totalRows, "处理块 " & chunkIndex & "..."
        
        ' 处理当前块
        Dim chunkRange As Range
        Set chunkRange = targetRange.Rows(startRow & ":" & endRow)
        
        If Not ProcessChunk(chunkRange, targetRange.Columns.Count) Then
            ProcessInChunks = False
            Exit Function
        End If
        
        g_ChunksProcessed = g_ChunksProcessed + 1
        
        ' 检查中断
        If CheckForCancel() Then
            g_CancelOperation = True
            ProcessInChunks = False
            Exit Function
        End If
        
        ' 定期释放内存
        If chunkIndex Mod 10 = 0 Then
            DoEvents
            CompactCache
        End If
    Next startRow
    
    ProcessInChunks = True
    Exit Function
    
ErrorHandler:
    ProcessInChunks = False
End Function

'--------------------------------------------------
' 处理单个块
'--------------------------------------------------
Private Function ProcessChunk(chunkRange As Range, totalColumns As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' 读取数据
    Dim dataArray As Variant
    dataArray = chunkRange.Value2
    
    ' 分析数据
    Dim columnAnalyses() As ColumnAnalysisData
    ReDim columnAnalyses(1 To totalColumns)
    
    Dim col As Long
    For col = 1 To totalColumns
        columnAnalyses(col) = AnalyzeColumnEnhanced(dataArray, col, chunkRange.Rows.Count, chunkRange.Columns(col))
    Next col
    
    ' 应用优化
    ApplyOptimizationToChunk chunkRange, columnAnalyses
    
    ProcessChunk = True
    Exit Function
    
ErrorHandler:
    ProcessChunk = False
End Function

'--------------------------------------------------
' 普通处理（不分块）
'--------------------------------------------------
Private Function ProcessNormal(targetRange As Range) As Boolean
    On Error GoTo ErrorHandler
    
    ShowProgress 0, 100, "正在分析数据..."
    
    ' 读取数据
    Dim dataArray As Variant
    dataArray = SafeReadRangeToArray(targetRange)
    
    ' 检查表头
    Dim hasHeader As Boolean
    If g_Config.SmartHeaderDetection And targetRange.Rows.Count > 1 Then
        hasHeader = IsHeaderRow(targetRange.Rows(1), targetRange.Rows(2))
    End If
    
    ' 分析列
    Dim columnAnalyses() As ColumnAnalysisData
    columnAnalyses = AnalyzeAllColumns(dataArray, targetRange)
    
    If g_CancelOperation Then
        ProcessNormal = False
        Exit Function
    End If
    
    ShowProgress 50, 100, "正在应用优化..."
    
    ' 应用优化
    ApplyColumnWidthOptimization targetRange, columnAnalyses
    ApplyAlignmentOptimizationWithHeader targetRange, columnAnalyses, hasHeader
    ApplyWrapAndRowHeight targetRange, columnAnalyses
    
    ProcessNormal = True
    Exit Function
    
ErrorHandler:
    ProcessNormal = False
End Function

'==================================================
' 缓存管理
'==================================================

'--------------------------------------------------
' 初始化缓存
'--------------------------------------------------
Private Sub InitializeCache()
    If Not g_Config.UseCache Then Exit Sub
    
    ReDim g_WidthCache(1 To CACHE_SIZE)
    g_CacheSize = 0
    g_CacheHits = 0
End Sub

'--------------------------------------------------
' 清空缓存
'--------------------------------------------------
Private Sub ClearCache()
    If Not g_Config.UseCache Then Exit Sub
    
    Erase g_WidthCache
    g_CacheSize = 0
    g_CacheHits = 0
End Sub

'--------------------------------------------------
' 压缩缓存（移除低频项）
'--------------------------------------------------
Private Sub CompactCache()
    If Not g_Config.UseCache Or g_CacheSize < CACHE_SIZE Then Exit Sub
    
    ' 按命中次数排序，保留前50%
    Dim i As Long, j As Long
    Dim temp As CellWidthCache
    
    ' 简单冒泡排序
    For i = 1 To g_CacheSize - 1
        For j = i + 1 To g_CacheSize
            If g_WidthCache(i).Hits < g_WidthCache(j).Hits Then
                temp = g_WidthCache(i)
                g_WidthCache(i) = g_WidthCache(j)
                g_WidthCache(j) = temp
            End If
        Next j
    Next i
    
    ' 保留前50%
    g_CacheSize = g_CacheSize \ 2
    ReDim Preserve g_WidthCache(1 To CACHE_SIZE)
End Sub

'--------------------------------------------------
' 获取缓存的宽度
'--------------------------------------------------
Private Function GetCachedWidth(content As String) As Double
    If Not g_Config.UseCache Then
        GetCachedWidth = CalculateTextWidth(content, 11)
        Exit Function
    End If
    
    ' 查找缓存
    Dim i As Long
    For i = 1 To g_CacheSize
        If g_WidthCache(i).Content = content Then
            g_WidthCache(i).Hits = g_WidthCache(i).Hits + 1
            g_CacheHits = g_CacheHits + 1
            GetCachedWidth = g_WidthCache(i).Width
            Exit Function
        End If
    Next i
    
    ' 计算并缓存
    Dim width As Double
    width = CalculateTextWidth(content, 11)
    
    ' 添加到缓存
    If g_CacheSize < CACHE_SIZE Then
        g_CacheSize = g_CacheSize + 1
        g_WidthCache(g_CacheSize).Content = content
        g_WidthCache(g_CacheSize).Width = width
        g_WidthCache(g_CacheSize).Hits = 1
    End If
    
    GetCachedWidth = width
End Function

'==================================================
' 增强的数据分析
'==================================================

'--------------------------------------------------
' 增强的列分析
'--------------------------------------------------
Private Function AnalyzeColumnEnhanced(dataArray As Variant, columnIndex As Long, rowCount As Long, columnRange As Range) As ColumnAnalysisData
    Dim analysis As ColumnAnalysisData
    Dim row As Long
    Dim cellValue As Variant
    Dim cellDataType As DataType
    Dim cellWidth As Double
    Dim maxWidth As Double
    
    analysis.ColumnIndex = columnIndex
    analysis.CellCount = 0
    analysis.HasMergedCells = HasMergedCells(columnRange)
    analysis.HasErrors = False
    maxWidth = 0
    
    ' 初始化标题相关字段
    analysis.IsHeaderColumn = False
    analysis.HeaderText = ""
    analysis.HeaderWidth = 0
    analysis.HeaderNeedWrap = False
    analysis.HeaderRowHeight = g_Config.HeaderMinHeight
    
    ' 检查是否为隐藏列，如果是隐藏列则跳过处理
    If columnRange.Hidden Then
        analysis.DataType = DataType.ShortText
        analysis.MaxContentWidth = columnRange.ColumnWidth ' 保持原始宽度
        analysis.OptimalWidth = columnRange.ColumnWidth
        analysis.NeedWrap = False
        AnalyzeColumnEnhanced = analysis
        Exit Function
    End If
    
    If analysis.HasMergedCells Then
        analysis.DataType = DataType.ShortText
        analysis.MaxContentWidth = 0
        analysis.OptimalWidth = 0
        analysis.NeedWrap = False
        AnalyzeColumnEnhanced = analysis
        Exit Function
    End If
    
    ' 分析标题（如果启用了智能表头检测）
    If g_Config.SmartHeaderDetection And rowCount > 0 Then
        Dim firstRowValue As Variant
        Dim secondRowValue As Variant
        
        If IsArray(dataArray) And UBound(dataArray, 2) >= columnIndex Then
            firstRowValue = dataArray(1, columnIndex)
            If rowCount > 1 Then secondRowValue = dataArray(2, columnIndex)
        End If
        
        ' 检查是否为标题行 - 放宽条件
        If Not IsEmpty(firstRowValue) And firstRowValue <> "" Then
            Dim headerRange As Range
            Set headerRange = columnRange.Cells(1, 1).Resize(1, 1)
            Dim dataRange As Range
            If rowCount > 1 Then Set dataRange = columnRange.Cells(2, 1).Resize(1, 1)
            
            ' 更宽松的标题识别：如果第一行包含中文或较长文本，倾向于认为是标题
            Dim isLikelyHeader As Boolean
            isLikelyHeader = False
            
            ' 条件1：传统的标题检测
            If Not dataRange Is Nothing And IsHeaderRow(headerRange, dataRange) Then
                isLikelyHeader = True
            End If
            
            ' 条件2：如果第一行文本较长且包含中文字符，可能是标题
            Dim headerText As String
            headerText = SafeGetCellValue(firstRowValue)
            If Len(headerText) >= 4 Then ' 长度>=4个字符
                ' 检查是否包含中文字符
                Dim i As Integer
                For i = 1 To Len(headerText)
                    Dim charCode As Integer
                    charCode = Asc(Mid(headerText, i, 1))
                    If charCode > 127 Or charCode < 0 Then ' 中文字符
                        isLikelyHeader = True
                        Exit For
                    End If
                Next i
            End If
            
            ' 条件3：如果第一行是纯文本且位置在第一行，默认作为标题处理
            If Not IsNumeric(firstRowValue) And Len(headerText) > 2 Then
                isLikelyHeader = True
            End If
            
            If isLikelyHeader Then
                analysis.IsHeaderColumn = True
                analysis.HeaderText = headerText
                analysis.HeaderWidth = AnalyzeHeaderWidth(analysis.HeaderText, g_Config.MaxColumnWidth)
                
                ' 判断标题是否需要换行
                Dim headerTextWidth As Double
                headerTextWidth = CalculateTextWidth(analysis.HeaderText, 12) ' 使用12号字体计算
                If headerTextWidth + g_Config.TextBuffer > g_Config.MaxColumnWidth Then
                    analysis.HeaderNeedWrap = True
                    analysis.HeaderRowHeight = CalculateHeaderRowHeight(analysis.HeaderText, g_Config.MaxColumnWidth)
                End If
            End If
        End If
    End If
    
    ' 分析数据内容（跳过标题行）
    Dim startRow As Long
    startRow = IIf(analysis.IsHeaderColumn, 2, 1)
    
    For row = startRow To rowCount
        If IsArray(dataArray) And UBound(dataArray, 2) >= columnIndex Then
            cellValue = dataArray(row, columnIndex)
        Else
            cellValue = dataArray
        End If
        
        If Not IsEmpty(cellValue) And cellValue <> "" Then
            analysis.CellCount = analysis.CellCount + 1
            
            ' 获取细化的数据类型
            cellDataType = GetEnhancedDataType(cellValue)
            analysis.TypeDistribution(cellDataType) = analysis.TypeDistribution(cellDataType) + 1
            
            If cellDataType = ErrorValue Then
                analysis.HasErrors = True
            End If
            
            If cellDataType <> ErrorValue Then
                cellWidth = GetCachedWidth(SafeGetCellValue(cellValue))
                If cellWidth > maxWidth Then
                    maxWidth = cellWidth
                End If
            End If
        End If
    Next row
    
    ' 确定主导数据类型
    analysis.DataType = DetermineColumnTypeEnhanced(analysis.TypeDistribution)
    
    ' 计算最优列宽（使用标题优先算法）
    analysis.MaxContentWidth = maxWidth
    Dim widthResult As WidthResult
    widthResult = CalculateOptimalWidthWithHeader(analysis)
    analysis.OptimalWidth = widthResult.FinalWidth
    analysis.NeedWrap = widthResult.NeedWrap
    
    AnalyzeColumnEnhanced = analysis
End Function

'--------------------------------------------------
' 获取增强的数据类型
'--------------------------------------------------
Private Function GetEnhancedDataType(cellValue As Variant) As DataType
    If IsError(cellValue) Then
        GetEnhancedDataType = ErrorValue
        Exit Function
    End If
    
    If IsEmpty(cellValue) Or cellValue = "" Then
        GetEnhancedDataType = EmptyCell
        Exit Function
    End If
    
    ' 布尔值检测
    If TypeName(cellValue) = "Boolean" Then
        GetEnhancedDataType = BooleanValue
        Exit Function
    End If
    
    ' 日期时间检测
    If IsDate(cellValue) Then
        Dim dateVal As Date
        dateVal = CDate(cellValue)
        
        If dateVal = Int(dateVal) Then
            GetEnhancedDataType = DateValue
        ElseIf dateVal < 1 Then
            GetEnhancedDataType = TimeValue
        Else
            GetEnhancedDataType = DateTimeValue
        End If
        Exit Function
    End If
    
    ' 数值检测
    If IsNumeric(cellValue) Then
        Dim numStr As String
        numStr = CStr(cellValue)
        
        If InStr(numStr, "%") > 0 Then
            GetEnhancedDataType = PercentageValue
        ElseIf InStr(numStr, "$") > 0 Or InStr(numStr, "¥") > 0 Or InStr(numStr, "€") > 0 Then
            GetEnhancedDataType = CurrencyValue
        ElseIf InStr(numStr, ".") > 0 Then
            GetEnhancedDataType = DecimalValue
        Else
            GetEnhancedDataType = IntegerValue
        End If
        Exit Function
    End If
    
    ' 文本类型细分
    Dim textLen As Long
    textLen = Len(CStr(cellValue))
    
    If textLen <= 10 Then
        GetEnhancedDataType = DataType.ShortText
    ElseIf textLen <= 50 Then
        GetEnhancedDataType = DataType.MediumText
    Else
        GetEnhancedDataType = DataType.LongText
    End If
End Function

'--------------------------------------------------
' 确定增强的列类型
'--------------------------------------------------
Private Function DetermineColumnTypeEnhanced(typeDistribution() As Long) As DataType
    Dim maxCount As Long
    Dim dominantType As DataType
    Dim i As Long
    
    maxCount = 0
    dominantType = DataType.ShortText
    
    ' 找出主导类型
    For i = 1 To 15
        If i <> EmptyCell And i <> ErrorValue Then
            If typeDistribution(i) > maxCount Then
                maxCount = typeDistribution(i)
                dominantType = i
            End If
        End If
    Next i
    
    ' 特殊规则
    ' 如果有长文本，整列按长文本处理
    If typeDistribution(DataType.LongText) > 0 Then
        dominantType = DataType.LongText
    ' 如果混合了文本和数值，按混合内容处理
    ElseIf (typeDistribution(DataType.ShortText) + typeDistribution(DataType.MediumText) > 0) And _
           (typeDistribution(IntegerValue) + typeDistribution(DecimalValue) > 0) Then
        dominantType = MixedContent
    End If
    
    DetermineColumnTypeEnhanced = dominantType
End Function

'--------------------------------------------------
' 计算增强的最优宽度
'--------------------------------------------------
Private Function CalculateOptimalWidthEnhanced(contentWidth As Double, dataType As DataType) As WidthResult
    Dim result As WidthResult
    Dim buffer As Double
    Dim calculatedWidth As Double
    
    If Not g_ConfigInitialized Then
        InitializeDefaultConfig
    End If
    
    ' 根据数据类型确定缓冲区
    Select Case dataType
        Case DataType.ShortText, DataType.MediumText
            buffer = g_Config.TextBuffer
        Case DataType.LongText
            buffer = g_Config.TextBuffer * 1.5
        Case IntegerValue, DecimalValue
            buffer = g_Config.NumericBuffer
        Case CurrencyValue, PercentageValue
            buffer = g_Config.NumericBuffer * 1.2
        Case DateValue, TimeValue, DateTimeValue
            buffer = g_Config.NumericBuffer
        Case MixedContent
            buffer = g_Config.TextBuffer * 1.2
        Case Else
            buffer = g_Config.TextBuffer
    End Select
    
    calculatedWidth = contentWidth + buffer
    result.OriginalWidth = calculatedWidth
    
    ' 应用边界控制
    If calculatedWidth < g_Config.MinColumnWidth Then
        result.FinalWidth = g_Config.MinColumnWidth
        result.NeedWrap = False
    ElseIf calculatedWidth >= g_Config.MaxColumnWidth Then
        result.FinalWidth = g_Config.MaxColumnWidth
        result.NeedWrap = (dataType = DataType.LongText Or dataType = DataType.MediumText)
    Else
        result.FinalWidth = calculatedWidth
        result.NeedWrap = False
    End If
    
    CalculateOptimalWidthEnhanced = result
End Function

'==================================================
' 错误处理增强
'==================================================

'--------------------------------------------------
' 错误分类
'--------------------------------------------------
Private Function ClassifyError(errorCode As Long, errorDesc As String) As ErrorInfo
    Dim info As ErrorInfo
    
    info.Code = errorCode
    info.Description = errorDesc
    
    Select Case errorCode
        Case 1004 ' 应用程序定义或对象定义错误
            info.Level = Fatal
            info.Action = "终止操作"
        Case 13 ' 类型不匹配
            info.Level = Severe
            info.Action = "跳过当前项"
        Case 18 ' 用户中断
            info.Level = Info
            info.Action = "取消操作"
        Case 6 ' 溢出
            info.Level = Warning
            info.Action = "使用默认值"
        Case Else
            info.Level = Warning
            info.Action = "记录并继续"
    End Select
    
    ClassifyError = info
End Function

'--------------------------------------------------
' 按级别处理错误
'--------------------------------------------------
Private Sub HandleErrorByLevel(errorInfo As ErrorInfo)
    Dim message As String
    
    Select Case errorInfo.Level
        Case Fatal
            message = "致命错误：" & errorInfo.Description & vbCrLf
            message = message & "操作已终止"
            MsgBox message, vbCritical, "Excel布局优化系统"
            
        Case Severe
            message = "严重错误：" & errorInfo.Description & vbCrLf
            message = message & "将" & errorInfo.Action
            MsgBox message, vbExclamation, "Excel布局优化系统"
            
        Case Warning
            ' 记录警告，不中断
            Debug.Print "警告: " & errorInfo.Description
            
        Case Info
            ' 信息级别，静默处理
            Debug.Print "信息: " & errorInfo.Description
    End Select
End Sub

'==================================================
' 配置管理增强
'==================================================

'--------------------------------------------------
' 初始化默认配置
'--------------------------------------------------
Private Sub InitializeDefaultConfig()
    With g_Config
        .MinColumnWidth = DEFAULT_MIN_COLUMN_WIDTH
        .MaxColumnWidth = DEFAULT_MAX_COLUMN_WIDTH
        .TextBuffer = 2.0  ' 从3.5减少到2.0
        .NumericBuffer = NUMERIC_BUFFER_CHARS
        .WrapThreshold = DEFAULT_MAX_COLUMN_WIDTH
        .SmartHeaderDetection = True
        .ShowPreview = True
        .AutoSave = True
        .UseCache = True
        .ChunkProcessing = True
        ' 标题相关新增配置
        .HeaderPriority = True      ' 启用标题优先模式
        .HeaderMaxWrapLines = 3     ' 标题最大换行3行
        .HeaderMinHeight = 15       ' 标题最小行高
        ' 超长文本处理配置（新增）
        .ExtremeTextWidth = EXTREME_TEXT_WIDTH       ' 极长文本固定宽度
        .LongTextThreshold = LONG_TEXT_THRESHOLD     ' 长文本阈值
        .SmartLineBreak = True                       ' 启用智能断行
        .MaxWrapLines = 3                            ' 最大换行行数（限制为3行）
        .LongTextExtendThreshold = LONG_TEXT_THRESHOLD ' 长文本扩展阈值
    End With
    g_ConfigInitialized = True
End Sub

'--------------------------------------------------
' 保存配置到工作簿
'--------------------------------------------------
Private Sub SaveConfigToWorkbook()
    On Error Resume Next
    
    Dim props As Object
    Set props = ThisWorkbook.CustomDocumentProperties
    
    ' 删除旧配置
    props("ExcelOptimizer_Config").Delete
    
    ' 保存新配置
    Dim configStr As String
    With g_Config
        configStr = .MinColumnWidth & "|" & .MaxColumnWidth & "|" & _
                   .TextBuffer & "|" & .NumericBuffer & "|" & _
                   .WrapThreshold & "|" & IIf(.SmartHeaderDetection, "1", "0") & "|" & _
                   IIf(.UseCache, "1", "0") & "|" & IIf(.ChunkProcessing, "1", "0")
    End With
    
    props.Add Name:="ExcelOptimizer_Config", _
              LinkToContent:=False, _
              Type:=msoPropertyTypeString, _
              Value:=configStr
End Sub

'--------------------------------------------------
' 从工作簿加载配置
'--------------------------------------------------
Private Sub LoadConfigFromWorkbook()
    On Error Resume Next
    
    Dim configStr As String
    configStr = ThisWorkbook.CustomDocumentProperties("ExcelOptimizer_Config").Value
    
    If configStr <> "" Then
        Dim parts() As String
        parts = Split(configStr, "|")
        
        If UBound(parts) >= 5 Then
            With g_Config
                .MinColumnWidth = CDbl(parts(0))
                .MaxColumnWidth = CDbl(parts(1))
                .TextBuffer = CDbl(parts(2))
                .NumericBuffer = CDbl(parts(3))
                .WrapThreshold = CDbl(parts(4))
                .SmartHeaderDetection = (parts(5) = "1")
                If UBound(parts) >= 7 Then
                    .UseCache = (parts(6) = "1")
                    .ChunkProcessing = (parts(7) = "1")
                End If
            End With
        End If
    End If
End Sub

'==================================================
' 验证增强
'==================================================

'--------------------------------------------------
' 增强的选择验证
'--------------------------------------------------
Private Function ValidateSelectionEnhanced(selectedRange As Range) As Boolean
    ValidateSelectionEnhanced = False
    
    ' 基础验证
    If selectedRange Is Nothing Then
        ShowErrorMessage "请先选择需要优化的区域", Warning
        Exit Function
    End If
    
    ' 工作表保护检查
    If selectedRange.Worksheet.ProtectContents Then
        ShowErrorMessage "工作表受保护，无法进行优化", Fatal
        Exit Function
    End If
    
    ' 大小检查
    Dim cellCount As Long
    cellCount = selectedRange.Cells.Count
    
    If cellCount > MAX_CELLS_LIMIT Then
        Dim response As VbMsgBoxResult
        Dim prompt As String
        prompt = "选择区域包含 " & Format(cellCount, "#,##0") & " 个单元格" & vbCrLf
        prompt = prompt & "处理可能需要较长时间" & vbCrLf
        prompt = prompt & "建议启用分块处理，是否继续？"
        response = MsgBox(prompt, vbYesNo + vbQuestion, "Excel布局优化系统")
        If response = vbNo Then Exit Function
        
        ' 自动启用分块处理
        g_Config.ChunkProcessing = True
    End If
    
    ' 合并单元格检查
    If HasMergedCells(selectedRange) Then
        Dim mergeResponse As VbMsgBoxResult
        mergeResponse = MsgBox("检测到合并单元格，这些区域将被跳过。是否继续？", _
                                vbYesNo + vbQuestion, "Excel布局优化系统")
        If mergeResponse = vbNo Then Exit Function
    End If
    
    ValidateSelectionEnhanced = True
End Function

'--------------------------------------------------
' 显示错误消息
'--------------------------------------------------
Private Sub ShowErrorMessage(message As String, level As ErrorLevel)
    Dim icon As VbMsgBoxStyle
    
    Select Case level
        Case Fatal
            icon = vbCritical
        Case Severe
            icon = vbExclamation
        Case Warning
            icon = vbInformation
        Case Else
            icon = vbInformation
    End Select
    
    MsgBox message, icon, "Excel布局优化系统"
End Sub

'==================================================
' 统计和报告增强
'==================================================

'--------------------------------------------------
' 生成增强的统计信息
'--------------------------------------------------
Private Function GenerateEnhancedStatistics(targetRange As Range, processingTime As Double) As OptimizationStats
    Dim stats As OptimizationStats
    
    stats.TotalColumns = targetRange.Columns.Count
    stats.ProcessingTime = processingTime
    stats.CacheHits = g_CacheHits
    stats.ChunksProcessed = g_ChunksProcessed
    
    ' 其他统计信息通过遍历获取
    Dim col As Long
    For col = 1 To stats.TotalColumns
        If targetRange.Columns(col).Hidden = False Then
            stats.AdjustedColumns = stats.AdjustedColumns + 1
            
            If targetRange.Columns(col).WrapText Then
                stats.WrapEnabledColumns = stats.WrapEnabledColumns + 1
            End If
        Else
            stats.SkippedColumns = stats.SkippedColumns + 1
        End If
    Next col
    
    ' 错误单元格统计
    Dim errorCount As Long
    On Error Resume Next
    errorCount = targetRange.SpecialCells(xlCellTypeFormulas, xlErrors).Count
    On Error GoTo 0
    stats.ErrorCount = errorCount
    
    GenerateEnhancedStatistics = stats
End Function

'--------------------------------------------------
' 显示增强的完成消息
'--------------------------------------------------
Private Sub ShowCompletionMessageEnhanced(stats As OptimizationStats)
    Dim message As String
    
    message = "优化完成！" & vbCrLf & vbCrLf
    message = message & "【处理统计】" & vbCrLf
    message = message & "• 处理列数：" & stats.TotalColumns & " 列" & vbCrLf
    message = message & "• 调整列数：" & stats.AdjustedColumns & " 列" & vbCrLf
    message = message & "• 启用换行：" & stats.WrapEnabledColumns & " 列" & vbCrLf
    
    If stats.SkippedColumns > 0 Then
        message = message & "• 跳过列数：" & stats.SkippedColumns & " 列" & vbCrLf
    End If
    
    message = message & vbCrLf & "【性能指标】" & vbCrLf
    message = message & "• 处理时间：" & Format(stats.ProcessingTime, "0.00") & " 秒" & vbCrLf
    
    If g_Config.UseCache Then
        message = message & "• 缓存命中：" & stats.CacheHits & " 次" & vbCrLf
    End If
    
    If g_Config.ChunkProcessing And stats.ChunksProcessed > 0 Then
        message = message & "• 处理块数：" & stats.ChunksProcessed & " 块" & vbCrLf
    End If
    
    If stats.ErrorCount > 0 Then
        message = message & vbCrLf & "【警告】" & vbCrLf
        message = message & "• 错误单元格：" & stats.ErrorCount & " 个" & vbCrLf
    End If
    
    message = message & vbCrLf & "提示：可使用 UndoLastOptimization 撤销本次操作"
    
    MsgBox message, vbInformation, "Excel布局优化系统"
End Sub

'==================================================
' 测试套件
'==================================================

'--------------------------------------------------
' 运行测试套件
'--------------------------------------------------
Public Sub RunTestSuite()
    Debug.Print "=" & String(50, "=")
    Debug.Print "Excel布局优化系统 - 测试套件"
    Debug.Print "开始时间: " & Now
    Debug.Print "=" & String(50, "=")
    
    Dim passCount As Long, failCount As Long
    
    ' 测试1：数据类型识别
    If TestDataTypeDetection() Then
        passCount = passCount + 1
        Debug.Print "✓ 数据类型识别测试通过"
    Else
        failCount = failCount + 1
        Debug.Print "✗ 数据类型识别测试失败"
    End If
    
    ' 测试2：列宽计算
    If TestColumnWidthCalculation() Then
        passCount = passCount + 1
        Debug.Print "✓ 列宽计算测试通过"
    Else
        failCount = failCount + 1
        Debug.Print "✗ 列宽计算测试失败"
    End If
    
    ' 测试3：缓存机制
    If TestCacheMechanism() Then
        passCount = passCount + 1
        Debug.Print "✓ 缓存机制测试通过"
    Else
        failCount = failCount + 1
        Debug.Print "✗ 缓存机制测试失败"
    End If
    
    ' 测试4：配置管理
    If TestConfigManagement() Then
        passCount = passCount + 1
        Debug.Print "✓ 配置管理测试通过"
    Else
        failCount = failCount + 1
        Debug.Print "✗ 配置管理测试失败"
    End If
    
    ' 测试5：标题优先功能（新增）
    If TestHeaderPriorityCalculation() Then
        passCount = passCount + 1
        Debug.Print "✓ 标题优先计算测试通过"
    Else
        failCount = failCount + 1
        Debug.Print "✗ 标题优先计算测试失败"
    End If
    
    ' 测试6：超长文本处理功能（新增）
    If TestExtremeTextProcessing() Then
        passCount = passCount + 1
        Debug.Print "✓ 超长文本处理测试通过"
    Else
        failCount = failCount + 1
        Debug.Print "✗ 超长文本处理测试失败"
    End If
    
    ' 测试7：安全数组读取功能
    If TestSafeReadRangeToArray() Then
        passCount = passCount + 1
        Debug.Print "✓ 安全数组读取测试通过"
    Else
        failCount = failCount + 1
        Debug.Print "✗ 安全数组读取测试失败"
    End If
    
    Debug.Print "=" & String(50, "=")
    Debug.Print "测试完成: 通过 " & passCount & " | 失败 " & failCount
    Debug.Print "结束时间: " & Now
    Debug.Print "=" & String(50, "=")
End Sub

'--------------------------------------------------
' 测试数据类型识别
'--------------------------------------------------
Private Function TestDataTypeDetection() As Boolean
    On Error GoTo TestFailed
    
    ' 测试各种数据类型
    Debug.Assert GetEnhancedDataType("Hello") = DataType.ShortText
    Debug.Assert GetEnhancedDataType("这是一段很长的文本内容用于测试长文本识别功能") = DataType.LongText
    Debug.Assert GetEnhancedDataType(123) = IntegerValue
    Debug.Assert GetEnhancedDataType(123.45) = DecimalValue
    Debug.Assert GetEnhancedDataType("50%") = PercentageValue
    Debug.Assert GetEnhancedDataType("$100") = CurrencyValue
    Debug.Assert GetEnhancedDataType(#1/1/2024#) = DateValue
    Debug.Assert GetEnhancedDataType(True) = BooleanValue
    
    TestDataTypeDetection = True
    Exit Function
    
TestFailed:
    TestDataTypeDetection = False
End Function

'--------------------------------------------------
' 测试列宽计算
'--------------------------------------------------
Private Function TestColumnWidthCalculation() As Boolean
    On Error GoTo TestFailed
    
    ' 测试不同文本的宽度计算
    Dim width1 As Double, width2 As Double, width3 As Double
    
    width1 = CalculateTextWidth("ABC", 11)
    width2 = CalculateTextWidth("中文测试", 11)
    width3 = CalculateTextWidth("123456", 11)
    
    Debug.Assert width1 > 0
    Debug.Assert width2 > width1  ' 中文应该更宽
    Debug.Assert width3 > 0
    
    TestColumnWidthCalculation = True
    Exit Function
    
TestFailed:
    TestColumnWidthCalculation = False
End Function

'--------------------------------------------------
' 测试缓存机制
'--------------------------------------------------
Private Function TestCacheMechanism() As Boolean
    On Error GoTo TestFailed
    
    ' 初始化缓存
    InitializeCache
    g_Config.UseCache = True
    
    ' 第一次调用（未命中）
    Dim width1 As Double
    width1 = GetCachedWidth("TestContent")
    
    ' 第二次调用（应该命中）
    Dim width2 As Double
    width2 = GetCachedWidth("TestContent")
    
    Debug.Assert width1 = width2
    Debug.Assert g_CacheHits > 0
    
    ClearCache
    TestCacheMechanism = True
    Exit Function
    
TestFailed:
    ClearCache
    TestCacheMechanism = False
End Function

'--------------------------------------------------
' 测试配置管理
'--------------------------------------------------
Private Function TestConfigManagement() As Boolean
    On Error GoTo TestFailed
    
    ' 保存当前配置
    Dim originalConfig As OptimizationConfig
    originalConfig = g_Config
    
    ' 修改配置
    g_Config.MaxColumnWidth = 75
    g_Config.UseCache = False
    
    ' 保存和加载
    SaveConfigToWorkbook
    
    ' 重置并重新加载
    InitializeDefaultConfig
    LoadConfigFromWorkbook
    
    ' 验证
    Debug.Assert g_Config.MaxColumnWidth = 75
    Debug.Assert g_Config.UseCache = False
    
    ' 恢复原始配置
    g_Config = originalConfig
    SaveConfigToWorkbook
    
    TestConfigManagement = True
    Exit Function
    
TestFailed:
    g_Config = originalConfig
    TestConfigManagement = False
End Function

'--------------------------------------------------
' 测试标题优先计算功能（新增）
'--------------------------------------------------
Private Function TestHeaderPriorityCalculation() As Boolean
    On Error GoTo TestFailed
    
    ' 保存原始配置
    Dim originalConfig As OptimizationConfig
    originalConfig = g_Config
    
    ' 设置测试配置
    g_Config.HeaderPriority = True
    g_Config.MaxColumnWidth = 50
    g_Config.TextBuffer = 2
    g_Config.HeaderMinHeight = 15
    g_Config.HeaderMaxWrapLines = 3
    
    ' 测试短标题（不需要换行）
    Dim shortHeaderWidth As Double
    shortHeaderWidth = AnalyzeHeaderWidth("姓名", 50)
    Debug.Assert shortHeaderWidth > 0 And shortHeaderWidth <= 50
    
    ' 测试长标题（需要换行）
    Dim longHeaderWidth As Double
    longHeaderWidth = AnalyzeHeaderWidth("这是一个非常长的标题用于测试换行功能是否正常工作", 30)
    Debug.Assert longHeaderWidth = 30 ' 应该返回最大宽度
    
    ' 测试行高计算
    Dim rowHeight As Double
    rowHeight = CalculateHeaderRowHeight("很长的标题文本需要换行显示", 20)
    Debug.Assert rowHeight >= g_Config.HeaderMinHeight
    
    ' 测试标题优先宽度计算
    Dim analysis As ColumnAnalysisData
    analysis.IsHeaderColumn = True
    analysis.HeaderText = "客户名称"
    analysis.MaxContentWidth = 15
    analysis.DataType = DataType.ShortText
    
    Dim result As WidthResult
    result = CalculateOptimalWidthWithHeader(analysis)
    Debug.Assert result.FinalWidth > 0
    
    ' 测试标题宽度大于数据宽度的情况
    analysis.HeaderText = "非常长的标题文本"
    analysis.MaxContentWidth = 5
    result = CalculateOptimalWidthWithHeader(analysis)
    Debug.Assert result.FinalWidth >= analysis.MaxContentWidth
    
    ' 恢复原始配置
    g_Config = originalConfig
    
    TestHeaderPriorityCalculation = True
    Exit Function
    
TestFailed:
    g_Config = originalConfig
    TestHeaderPriorityCalculation = False
End Function

'--------------------------------------------------
' 测试超长文本处理功能（新增）
'--------------------------------------------------
Private Function TestExtremeTextProcessing() As Boolean
    On Error GoTo TestFailed
    
    ' 保存原始配置
    Dim originalConfig As OptimizationConfig
    originalConfig = g_Config
    
    ' 测试1：文本长度分类
    Dim shortText As String, longText As String, extremeText As String
    shortText = "短文本"
    longText = "这是一个比较长的文本内容，用来测试长文本的处理效果和分类准确性，确保系统能够正确识别不同长度的文本"
    extremeText = "这是一个极长的文本内容，专门设计用来测试系统在处理极端长度文本时的表现，包括但不限于：智能换行处理、行高自动调整、列宽优化计算、文本截断保护、格式保持、可读性优化、性能控制等多个方面的功能，确保系统能够在各种极端情况下都能够稳定运行"
    
    If ClassifyTextLength(shortText) <> TextLengthCategory.ShortText Then GoTo TestFailed
    If ClassifyTextLength(longText) <> TextLengthCategory.LongText Then GoTo TestFailed
    If ClassifyTextLength(extremeText) <> TextLengthCategory.ExtremeText Then GoTo TestFailed
    
    ' 测试2：超长文本宽度计算
    Dim calculatedWidth As Double
    calculatedWidth = CalculateExtremeTextWidth(extremeText)
    If calculatedWidth <> g_Config.ExtremeTextWidth Then GoTo TestFailed
    
    ' 测试3：智能换行布局计算
    Dim layout As WrapLayout
    layout = CalculateWrapLayout(extremeText, 100)
    If Not layout.NeedWrap Then GoTo TestFailed
    If layout.TotalLines <= 1 Then GoTo TestFailed
    
    ' 测试4：断行点查找
    Dim breaks As Collection
    Set breaks = FindBreakPoints("测试，文本；内容：换行！效果？验证。")
    If breaks.Count = 0 Then GoTo TestFailed
    
    ' 测试5：行高计算
    Dim rowHeight As Double
    rowHeight = CalculateOptimalRowHeight(extremeText, 120)
    If rowHeight <= MIN_ROW_HEIGHT Then GoTo TestFailed
    
    ' 恢复原始配置
    g_Config = originalConfig
    
    TestExtremeTextProcessing = True
    Exit Function
    
TestFailed:
    g_Config = originalConfig
    TestExtremeTextProcessing = False
End Function

'--------------------------------------------------
' 测试安全数组读取功能
'--------------------------------------------------
Private Function TestSafeReadRangeToArray() As Boolean
    On Error GoTo TestFailed
    
    ' 创建测试区域
    Dim testWs As Worksheet
    Set testWs = ActiveSheet
    
    ' 保存原始值
    Dim originalA1 As Variant, originalB1 As Variant
    originalA1 = testWs.Range("A1").Value
    originalB1 = testWs.Range("B1").Value
    
    ' 测试单个单元格
    testWs.Range("A1").Value = "测试值"
    Dim singleResult As Variant
    singleResult = SafeReadRangeToArray(testWs.Range("A1"))
    Debug.Assert IsArray(singleResult)
    Debug.Assert UBound(singleResult, 1) = 1 And UBound(singleResult, 2) = 1
    Debug.Assert singleResult(1, 1) = "测试值"
    
    ' 测试多单元格区域
    testWs.Range("A1").Value = "标题1"
    testWs.Range("B1").Value = "标题2"
    Dim multiResult As Variant
    multiResult = SafeReadRangeToArray(testWs.Range("A1:B1"))
    Debug.Assert IsArray(multiResult)
    Debug.Assert UBound(multiResult, 2) = 2
    Debug.Assert multiResult(1, 1) = "标题1"
    Debug.Assert multiResult(1, 2) = "标题2"
    
    ' 测试空值处理
    testWs.Range("A1").ClearContents
    Dim emptyResult As Variant
    emptyResult = SafeReadRangeToArray(testWs.Range("A1"))
    Debug.Assert IsArray(emptyResult)
    
    ' 恢复原始值
    testWs.Range("A1").Value = originalA1
    testWs.Range("B1").Value = originalB1
    
    TestSafeReadRangeToArray = True
    Exit Function
    
TestFailed:
    ' 恢复原始值
    On Error Resume Next
    testWs.Range("A1").Value = originalA1
    testWs.Range("B1").Value = originalB1
    TestSafeReadRangeToArray = False
End Function

'==================================================
' 配置和用户交互函数
'==================================================

'--------------------------------------------------
' 获取用户配置
'--------------------------------------------------
Private Function GetUserConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    Dim response As String
    Dim prompt As String
    
    ' 简单配置模式 - 询问关键参数
    prompt = "设置最大列宽（字符单位）" & vbCrLf & _
             "范围: 30-100，默认: " & g_Config.MaxColumnWidth & vbCrLf & _
             "直接按Enter使用默认值，按取消退出配置"
    
    response = InputBox(prompt, "Excel布局优化配置", CStr(g_Config.MaxColumnWidth))
    
    ' 用户取消
    If StrPtr(response) = 0 Then
        GetUserConfiguration = False
        Exit Function
    End If
    
    ' 用户按Enter使用默认值
    If response = "" Then
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
            MsgBox "请输入30-100之间的数值", vbExclamation, "输入错误"
            GetUserConfiguration = False
            Exit Function
        End If
    Else
        MsgBox "请输入有效的数字", vbExclamation, "输入错误"
        GetUserConfiguration = False
        Exit Function
    End If
    
    ' 询问是否显示预览
    Dim previewResponse As VbMsgBoxResult
    previewResponse = MsgBox("是否在优化前显示预览？", vbYesNo + vbQuestion, "预览设置")
    g_Config.ShowPreview = (previewResponse = vbYes)
    
    GetUserConfiguration = True
    Exit Function
    
ErrorHandler:
    GetUserConfiguration = False
End Function

'--------------------------------------------------
' 收集预览信息
'--------------------------------------------------
Private Function CollectPreviewInfo(targetRange As Range) As PreviewInfo
    On Error GoTo ErrorHandler
    
    Dim info As PreviewInfo
    
    With info
        .TotalColumns = targetRange.Columns.Count
        .AffectedCells = targetRange.Cells.Count
        .HasMergedCells = HasMergedCells(targetRange)
        .HasFormulas = HasFormulas(targetRange)
        
        ' 分析列宽变化
        Dim col As Long
        Dim maxWidth As Double, minWidth As Double
        minWidth = 999
        maxWidth = 0
        
        For col = 1 To .TotalColumns
            If Not targetRange.Columns(col).Hidden Then
                Dim currentWidth As Double
                currentWidth = targetRange.Columns(col).ColumnWidth
                
                If currentWidth < minWidth Then minWidth = currentWidth
                If currentWidth > maxWidth Then maxWidth = currentWidth
                
                ' 简单估算需要调整的列
                If currentWidth < g_Config.MinColumnWidth Or _
                   currentWidth > g_Config.MaxColumnWidth Then
                    .ColumnsToAdjust = .ColumnsToAdjust + 1
                End If
            End If
        Next col
        
        .MinWidth = minWidth
        .MaxWidth = maxWidth
        
        ' 估算需要换行的列（简化版）
        For col = 1 To .TotalColumns
            If targetRange.Columns(col).ColumnWidth > g_Config.WrapThreshold Then
                .ColumnsNeedWrap = .ColumnsNeedWrap + 1
            End If
        Next col
        
        ' 估算处理时间
        .EstimatedTime = (.AffectedCells / 10000) * 1.5
        If .EstimatedTime < 0.5 Then .EstimatedTime = 0.5
    End With
    
    CollectPreviewInfo = info
    Exit Function
    
ErrorHandler:
    ' 返回默认信息
    With info
        .TotalColumns = targetRange.Columns.Count
        .AffectedCells = targetRange.Cells.Count
        .EstimatedTime = 1
    End With
    CollectPreviewInfo = info
End Function

'--------------------------------------------------
' 显示预览对话框
'--------------------------------------------------
Private Function ShowPreviewDialog(info As PreviewInfo, targetRange As Range) As VbMsgBoxResult
    On Error GoTo ErrorHandler
    
    Dim message As String
    
    message = "布局优化预览" & vbCrLf & vbCrLf
    message = message & "优化区域: " & targetRange.Address & vbCrLf
    message = message & String(40, "-") & vbCrLf
    message = message & "• 总列数: " & info.TotalColumns & vbCrLf
    message = message & "• 影响单元格: " & Format(info.AffectedCells, "#,##0") & vbCrLf
    
    If info.ColumnsToAdjust > 0 Then
        message = message & "• 需调整: " & info.ColumnsToAdjust & " 列" & vbCrLf
    End If
    
    If info.ColumnsNeedWrap > 0 Then
        message = message & "• 可能需要换行: " & info.ColumnsNeedWrap & " 列" & vbCrLf
    End If
    
    If info.MinWidth > 0 And info.MaxWidth > 0 Then
        message = message & "• 当前列宽范围: " & Format(info.MinWidth, "0.0") & _
                  " - " & Format(info.MaxWidth, "0.0") & vbCrLf
    End If
    
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
    Exit Function
    
ErrorHandler:
    ShowPreviewDialog = vbNo
End Function

'--------------------------------------------------
' 保存撤销信息
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
' 恢复撤销信息
'--------------------------------------------------
Private Function RestoreFromUndo() As Boolean
    On Error GoTo ErrorHandler
    
    If Not g_HasUndoInfo Then
        RestoreFromUndo = False
        Exit Function
    End If
    
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

'--------------------------------------------------
' 安全读取范围到数组
'--------------------------------------------------
Private Function SafeReadRangeToArray(targetRange As Range) As Variant
    On Error GoTo ErrorHandler
    
    Dim result As Variant
    
    ' 处理单个单元格的情况
    If targetRange.Cells.Count = 1 Then
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = targetRange.Value
        SafeReadRangeToArray = result
        Exit Function
    End If
    
    ' 处理多个单元格
    result = targetRange.Value2
    
    ' 确保返回的是二维数组
    If Not IsArray(result) Then
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = targetRange.Value
    End If
    
    SafeReadRangeToArray = result
    Exit Function
    
ErrorHandler:
    ' 错误时返回空数组
    ReDim result(1 To targetRange.Rows.Count, 1 To targetRange.Columns.Count)
    SafeReadRangeToArray = result
End Function

'--------------------------------------------------
' 应用对齐优化（修正版，使用定义的辅助函数）
'--------------------------------------------------
Private Sub ApplyAlignmentOptimizationWithHeader(targetRange As Range, analyses() As ColumnAnalysisData, hasHeader As Boolean)
    ' 智能对齐优化，支持标题居中，但保护隐藏列
    On Error Resume Next
    
    Dim i As Long
    Dim col As Range
    
    For i = 1 To UBound(analyses)
        Set col = targetRange.Columns(i)
        
        ' 只处理可见列
        If Not col.Hidden Then
            ' 如果有表头且是第一行，标题居中
            If hasHeader And targetRange.Rows.Count > 1 Then
                ' 标题行（第一行）居中
                If Not targetRange.Rows(1).Hidden Then
                    col.Cells(1, 1).HorizontalAlignment = xlCenter
                    col.Cells(1, 1).VerticalAlignment = xlCenter
                End If
                
                ' 数据行根据类型对齐
                If targetRange.Rows.Count > 1 Then
                    Dim dataRange As Range
                    Set dataRange = col.Resize(targetRange.Rows.Count - 1, 1).Offset(1, 0)
                    
                    ' 过滤出可见行进行处理
                    Dim visibleDataRange As Range
                    Set visibleDataRange = GetVisibleRange(dataRange)
                    
                    If Not visibleDataRange Is Nothing Then
                        Select Case analyses(i).DataType
                            Case IntegerValue, DecimalValue, CurrencyValue, PercentageValue
                                visibleDataRange.HorizontalAlignment = xlRight
                            Case DateValue, TimeValue, DateTimeValue
                                visibleDataRange.HorizontalAlignment = xlCenter
                            Case Else
                                visibleDataRange.HorizontalAlignment = xlLeft
                        End Select
                        visibleDataRange.VerticalAlignment = xlTop
                    End If
                End If
            Else
                ' 没有表头，整列统一对齐
                Dim visibleCol As Range
                Set visibleCol = GetVisibleRange(col)
                
                If Not visibleCol Is Nothing Then
                    Select Case analyses(i).DataType
                        Case IntegerValue, DecimalValue, CurrencyValue, PercentageValue
                            visibleCol.HorizontalAlignment = xlRight
                        Case DateValue, TimeValue, DateTimeValue
                            visibleCol.HorizontalAlignment = xlCenter
                        Case Else
                            visibleCol.HorizontalAlignment = xlLeft
                    End Select
                    visibleCol.VerticalAlignment = xlTop
                End If
            End If
        End If
    Next i
    
    On Error GoTo 0
End Sub

'--------------------------------------------------
' 应用换行和行高调整（修正版）
'--------------------------------------------------
Private Sub ApplyWrapAndRowHeight(targetRange As Range, analyses() As ColumnAnalysisData)
    ' 应用换行和行高调整，特别关注标题行和超长文本
    On Error Resume Next
    
    Dim i As Long
    Dim hasHeaderAdjustment As Boolean
    Dim maxHeaderHeight As Double
    hasHeaderAdjustment = False
    maxHeaderHeight = 15 ' 默认最小行高
    
    ' 首先处理列的换行设置
    For i = 1 To UBound(analyses)
        ' 跳过隐藏列
        If targetRange.Columns(i).Hidden Then
            GoTo NextColumn
        End If
        
        ' 处理需要换行的列
        If analyses(i).NeedWrap And Not analyses(i).HasMergedCells Then
            ' 获取可见单元格
            Dim visibleColCells As Range
            Set visibleColCells = GetVisibleRange(targetRange.Columns(i))
            
            If Not visibleColCells Is Nothing Then
                visibleColCells.WrapText = True
            End If
        End If
        
        ' 处理标题换行（如果存在标题）
        If analyses(i).IsHeaderColumn And analyses(i).HeaderNeedWrap Then
            If Not targetRange.Rows(1).Hidden Then
                targetRange.Columns(i).Cells(1, 1).WrapText = True
                hasHeaderAdjustment = True
                If analyses(i).HeaderRowHeight > maxHeaderHeight Then
                    maxHeaderHeight = analyses(i).HeaderRowHeight
                End If
            End If
        End If
        
NextColumn:
    Next i
    
    ' 如果有标题需要调整，先设置标题行高（仅在标题行可见时）
    If hasHeaderAdjustment And Not targetRange.Rows(1).Hidden Then
        targetRange.Rows(1).RowHeight = maxHeaderHeight
    End If
    
    ' 自动调整所有可见行的行高（但保护已设置的标题行高）
    Dim originalFirstRowHeight As Double
    If hasHeaderAdjustment And Not targetRange.Rows(1).Hidden Then
        originalFirstRowHeight = targetRange.Rows(1).RowHeight
    End If
    
    ' 对可见的数据行进行自动调整
    If targetRange.Rows.Count > 1 Then
        Dim dataRows As Range
        Set dataRows = targetRange.Rows("2:" & targetRange.Rows.Count)
        
        ' 获取可见的数据行
        Dim visibleDataRows As Range
        Set visibleDataRows = GetVisibleRange(dataRows)
        
        If Not visibleDataRows Is Nothing Then
            visibleDataRows.AutoFit
        End If
    End If
    
    ' 恢复标题行高（如果被自动调整影响了）
    If hasHeaderAdjustment And Not targetRange.Rows(1).Hidden Then
        targetRange.Rows(1).RowHeight = originalFirstRowHeight
    End If
    
    On Error GoTo 0
End Sub

'--------------------------------------------------
' 计算文本宽度（核心算法）
'--------------------------------------------------
Private Function CalculateTextWidth(text As String, fontSize As Integer) As Double
    On Error GoTo ErrorHandler
    
    If Len(text) = 0 Then
        CalculateTextWidth = 0
        Exit Function
    End If
    
    ' 统计字符类型
    Dim charStats As CharCount
    charStats = CountCharacterTypes(text)
    
    ' 计算加权宽度
    Dim width As Double
    width = charStats.ChineseCount * CHINESE_CHAR_WIDTH_FACTOR + _
            charStats.EnglishCount * ENGLISH_CHAR_WIDTH_FACTOR + _
            charStats.NumberCount * NUMBER_CHAR_WIDTH_FACTOR + _
            charStats.OtherCount * OTHER_CHAR_WIDTH_FACTOR
    
    ' 字号调整（11号字体为基准）
    width = width * (fontSize / 11)
    
    CalculateTextWidth = width
    Exit Function
    
ErrorHandler:
    CalculateTextWidth = Len(text) * 0.7 ' 默认值
End Function

'--------------------------------------------------
' 统计字符类型
'--------------------------------------------------
Private Function CountCharacterTypes(text As String) As CharCount
    Dim stats As CharCount
    Dim i As Long
    Dim char As String
    Dim charCode As Integer
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        charCode = Asc(char)
        
        If charCode > 127 Or charCode < 0 Then
            ' 中文字符
            stats.ChineseCount = stats.ChineseCount + 1
        ElseIf (charCode >= 65 And charCode <= 90) Or _
               (charCode >= 97 And charCode <= 122) Then
            ' 英文字符
            stats.EnglishCount = stats.EnglishCount + 1
        ElseIf charCode >= 48 And charCode <= 57 Then
            ' 数字
            stats.NumberCount = stats.NumberCount + 1
        Else
            ' 其他字符
            stats.OtherCount = stats.OtherCount + 1
        End If
    Next i
    
    stats.TotalCount = Len(text)
    CountCharacterTypes = stats
End Function

'--------------------------------------------------
' 分析标题宽度
'--------------------------------------------------
Private Function AnalyzeHeaderWidth(headerText As String, maxWidth As Double) As Double
    On Error GoTo ErrorHandler
    
    If Len(headerText) = 0 Then
        AnalyzeHeaderWidth = g_Config.MinColumnWidth
       
        Exit Function
    End If
    
    ' 计算标题的基本宽度（包含缓冲）
    Dim baseWidth As Double
    baseWidth = CalculateTextWidth(headerText, 11) + g_Config.TextBuffer
    
    ' 如果标题宽度在限制范围内，直接返回
    If baseWidth <= maxWidth Then
        AnalyzeHeaderWidth = baseWidth
    Else
        ' 标题需要换行，返回最大宽度
        AnalyzeHeaderWidth = maxWidth
    End If
    
    Exit Function
    
ErrorHandler:
    AnalyzeHeaderWidth = g_Config.MinColumnWidth
End Function

'--------------------------------------------------
' 计算标题行高
'--------------------------------------------------
Private Function CalculateHeaderRowHeight(headerText As String, columnWidth As Double) As Double
    On Error GoTo ErrorHandler
    
    ' 计算需要的行数
    Dim textWidth As Double
    textWidth = CalculateTextWidth(headerText, 11)
    
    Dim linesNeeded As Long
    linesNeeded = Application.Max(1, Application.WorksheetFunction.Ceiling(textWidth / columnWidth, 1))
    
    ' 限制最大行数避免过度换行
    If linesNeeded > g_Config.HeaderMaxWrapLines Then
        linesNeeded = g_Config.HeaderMaxWrapLines
    End If
    
    ' 计算行高（每行约18像素包含间距）
    CalculateHeaderRowHeight = Application.Max(g_Config.HeaderMinHeight, linesNeeded * 18)
    
    Exit Function
    
ErrorHandler:
    CalculateHeaderRowHeight = g_Config.HeaderMinHeight
End Function

'--------------------------------------------------
' 标题优先的列宽计算
'--------------------------------------------------
Private Function CalculateOptimalWidthWithHeader(analysis As ColumnAnalysisData) As WidthResult
    Dim result As WidthResult
    On Error GoTo ErrorHandler
    
    ' 如果不是标题列或没有启用标题优先，使用原有逻辑
    If Not analysis.IsHeaderColumn Or Not g_Config.HeaderPriority Then
        result = CalculateOptimalWidthEnhanced(analysis.MaxContentWidth, analysis.DataType)
        CalculateOptimalWidthWithHeader = result
        Exit Function
    End If
    
    ' 标题优先模式：标题宽度 vs 数据宽度
    Dim headerRequiredWidth As Double
    Dim dataOptimalWidth As Double
    
    ' 计算标题需要的宽度
    headerRequiredWidth = AnalyzeHeaderWidth(analysis.HeaderText, g_Config.MaxColumnWidth)
    
    ' 计算数据内容的最优宽度
    dataOptimalWidth = analysis.MaxContentWidth + g_Config.TextBuffer
    If dataOptimalWidth < g_Config.MinColumnWidth Then
        dataOptimalWidth = g_Config.MinColumnWidth
    End If
    
    ' 取两者中的较大值作为最终宽度
    result.FinalWidth = Application.Max(headerRequiredWidth, dataOptimalWidth)
    
    ' 检查是否需要换行
    Dim headerTextWidth As Double
    headerTextWidth = CalculateTextWidth(analysis.HeaderText, 11)
    
    If headerTextWidth + g_Config.TextBuffer > g_Config.MaxColumnWidth Then
        result.NeedWrap = True
        result.FinalWidth = g_Config.MaxColumnWidth
    Else
        result.NeedWrap = False
    End If
    
    ' 应用最终的边界控制
    If result.FinalWidth > g_Config.MaxColumnWidth Then
        result.FinalWidth = g_Config.MaxColumnWidth
        result.NeedWrap = True
    ElseIf result.FinalWidth < g_Config.MinColumnWidth Then
        result.FinalWidth = g_Config.MinColumnWidth
    End If
    
    result.OriginalWidth = analysis.MaxContentWidth
    CalculateOptimalWidthWithHeader = result
    
    Exit Function
    
ErrorHandler:
    ' 错误情况下返回安全值
    result.FinalWidth = g_Config.MinColumnWidth
    result.NeedWrap = False
    result.OriginalWidth = 0
    CalculateOptimalWidthWithHeader = result
End Function

'--------------------------------------------------
' 文本长度分类
'--------------------------------------------------
Private Function ClassifyTextLength(text As String) As TextLengthCategory
    Dim length As Long
    length = Len(text)
    
    If length <= 20 Then
        ClassifyTextLength = TextLengthCategory.ShortText
    ElseIf length <= 50 Then
        ClassifyTextLength = TextLengthCategory.MediumText
    ElseIf length <= 100 Then
        ClassifyTextLength = TextLengthCategory.LongText
    ElseIf length <= 200 Then
        ClassifyTextLength = TextLengthCategory.VeryLongText
    Else
        ClassifyTextLength = TextLengthCategory.ExtremeText
    End If
End Function

'--------------------------------------------------
' 计算极长文本宽度
'--------------------------------------------------
Private Function CalculateExtremeTextWidth(text As String) As Double
    ' 对于极长文本，使用固定宽度
    CalculateExtremeTextWidth = g_Config.ExtremeTextWidth
End Function

'--------------------------------------------------
' 计算换行布局
'--------------------------------------------------
Private Function CalculateWrapLayout(text As String, columnWidth As Double) As WrapLayout
    Dim layout As WrapLayout
    
    ' 计算文本总宽度
    Dim textWidth As Double
    textWidth = CalculateTextWidth(text, 11)
    
    ' 计算需要的行数
    layout.TotalLines = Application.WorksheetFunction.Ceiling(textWidth / columnWidth, 1)
    
    ' 限制最大行数
    If layout.TotalLines > g_Config.MaxWrapLines Then
        layout.TotalLines = g_Config.MaxWrapLines
    End If
    
    ' 计算最优行高
    layout.OptimalRowHeight = Application.Max(MIN_ROW_HEIGHT, layout.TotalLines * 18)
    
    ' 是否需要换行
    layout.NeedWrap = (layout.TotalLines > 1)
    
    CalculateWrapLayout = layout
End Function

'--------------------------------------------------
' 查找断行点
'--------------------------------------------------
Private Function FindBreakPoints(text As String) As Collection
    Dim breaks As New Collection
    Dim i As Long
    Dim char As String
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        ' 检查标点符号
        If InStr("，。；：！？,;:!?、", char) > 0 Then
            breaks.Add i
        ' 检查空格
        ElseIf char = " " Then
            breaks.Add i
        End If
    Next i
    
    Set FindBreakPoints = breaks
End Function

'--------------------------------------------------
' 计算最优行高
'--------------------------------------------------
Private Function CalculateOptimalRowHeight(text As String, columnWidth As Double) As Double
    On Error GoTo ErrorHandler
    
    Dim baseHeight As Double
    baseHeight = MIN_ROW_HEIGHT
    
    ' 计算需要的行数
    Dim textWidth As Double
    textWidth = CalculateTextWidth(text, 11)
    
    Dim lines As Long
    lines = Application.WorksheetFunction.Ceiling(textWidth / (columnWidth * PIXELS_PER_CHAR_UNIT), 1)
    
    ' 限制最大行数
    If lines > 11 Then lines = 11
    
    ' 计算总高度（包含行间距）
    Dim totalHeight As Double
    totalHeight = baseHeight + (lines - 1) * 18
    
    ' 应用最大高度限制
    If totalHeight > MAX_ROW_HEIGHT Then totalHeight = MAX_ROW_HEIGHT
    
    CalculateOptimalRowHeight = totalHeight
    Exit Function
    
ErrorHandler:
    CalculateOptimalRowHeight = MIN_ROW_HEIGHT
End Function