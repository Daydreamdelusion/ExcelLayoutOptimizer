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
Private Const MAX_WRAP_LINES As Long = 10               ' 最大换行行数

' 列宽边界控制（像素）
Private Const MIN_COLUMN_WIDTH_PIXELS As Long = 50
Private Const MAX_COLUMN_WIDTH_PIXELS As Long = 300

' 缓冲区设置（像素）
Private Const TEXT_BUFFER_PIXELS As Long = 15
Private Const NUMERIC_BUFFER_PIXELS As Long = 12
Private Const DATE_BUFFER_PIXELS As Long = 12

' 缓冲区设置（字符单位）
Private Const TEXT_BUFFER_CHARS As Double = 3.5    ' 增加到3.5个字符单位以确保中文标题完整显示
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
            
            ' 条件2：如果第一行文本较长且包含中文字符，很可能是标题
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
        .TextBuffer = TEXT_BUFFER_CHARS
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
        .MaxWrapLines = MAX_WRAP_LINES               ' 最大换行行数
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
' 辅助函数
'==================================================

'--------------------------------------------------
' 分析所有列
'--------------------------------------------------
Private Function AnalyzeAllColumns(dataArray As Variant, targetRange As Range) As ColumnAnalysisData()
    Dim rowCount As Long, colCount As Long
    Dim col As Long
    
    If IsArray(dataArray) Then
        rowCount = UBound(dataArray, 1)
        colCount = UBound(dataArray, 2)
    Else
        rowCount = 1
        colCount = 1
    End If
    
    Dim analyses() As ColumnAnalysisData
    ReDim analyses(1 To colCount)
    
    For col = 1 To colCount
        ShowProgress 20 + (col - 1) * 30 / colCount, 100, "分析列 " & col & "/" & colCount
        
        If CheckForCancel() Then
            g_CancelOperation = True
            AnalyzeAllColumns = analyses
            Exit Function
        End If
        
        ' 检查列是否隐藏，如果隐藏则跳过分析，但保留默认结构
        If targetRange.Columns(col).Hidden Then
            ' 为隐藏列创建一个默认的分析结果，但不进行实际分析
            Dim defaultAnalysis As ColumnAnalysisData
            defaultAnalysis.OptimalWidth = targetRange.Columns(col).ColumnWidth ' 保持原始宽度
            defaultAnalysis.DataType = EmptyCell
            defaultAnalysis.IsHeaderColumn = False
            defaultAnalysis.HasMergedCells = False
            defaultAnalysis.NeedWrap = False
            defaultAnalysis.HeaderNeedWrap = False
            analyses(col) = defaultAnalysis
        Else
            ' 只分析可见列
            analyses(col) = AnalyzeColumnEnhanced(dataArray, col, rowCount, targetRange.Columns(col))
        End If
    Next col
    
    AnalyzeAllColumns = analyses
End Function

'--------------------------------------------------
' 应用优化到块
'--------------------------------------------------
Private Sub ApplyOptimizationToChunk(chunkRange As Range, columnAnalyses() As ColumnAnalysisData)
    Dim col As Long
    Dim hasHeaderRowAdjustment As Boolean
    hasHeaderRowAdjustment = False
    
    For col = 1 To UBound(columnAnalyses)
        ' 跳过隐藏列
        If Not chunkRange.Columns(col).Hidden And Not columnAnalyses(col).HasMergedCells And columnAnalyses(col).OptimalWidth > 0 Then
            ' 只在第一个块时设置列宽
            If chunkRange.Row = chunkRange.Parent.UsedRange.Row Then
                chunkRange.Columns(col).EntireColumn.ColumnWidth = columnAnalyses(col).OptimalWidth
            End If
            
            ' 设置对齐和换行
            If columnAnalyses(col).NeedWrap Then
                chunkRange.Columns(col).WrapText = True
            End If
            
            ' 处理标题行高调整
            If columnAnalyses(col).IsHeaderColumn And columnAnalyses(col).HeaderNeedWrap Then
                If chunkRange.Row = chunkRange.Parent.UsedRange.Row Then
                    ' 只在处理第一个块时调整标题行高
                    chunkRange.Columns(col).Cells(1, 1).WrapText = True
                    hasHeaderRowAdjustment = True
                End If
            End If
        End If
    Next col
    
    ' 如果有标题需要调整行高，统一设置第一行行高
    If hasHeaderRowAdjustment And chunkRange.Row = chunkRange.Parent.UsedRange.Row Then
        Dim maxHeaderHeight As Double
        maxHeaderHeight = g_Config.HeaderMinHeight
        
        ' 找出需要的最大行高
        For col = 1 To UBound(columnAnalyses)
            If columnAnalyses(col).IsHeaderColumn And columnAnalyses(col).HeaderNeedWrap Then
                If columnAnalyses(col).HeaderRowHeight > maxHeaderHeight Then
                    maxHeaderHeight = columnAnalyses(col).HeaderRowHeight
                End If
            End If
        Next col
        
        ' 设置第一行行高
        chunkRange.Rows(1).RowHeight = maxHeaderHeight
    End If
End Sub

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
    
    If total = 0 Then Exit Sub ' 防止除零错误
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
' 快捷键绑定
'==================================================

'--------------------------------------------------
' 安装快捷键
'--------------------------------------------------
Public Sub InstallShortcuts()
    On Error Resume Next
    Application.OnKey "^+L", "OptimizeLayout"
    Application.OnKey "^+Q", "QuickOptimize"
    Application.OnKey "^+H", "ShowSystemInfo"
    If Err.Number = 0 Then
        MsgBox "快捷键已安装成功！" & vbCrLf & vbCrLf & _
               "Ctrl+Shift+L - 布局优化" & vbCrLf & _
               "Ctrl+Shift+Q - 快速优化" & vbCrLf & _
               "Ctrl+Shift+H - 系统信息", _
               vbInformation, "快捷键安装"
    Else
        MsgBox "快捷键安装失败", vbCritical, "错误"
    End If
    On Error GoTo 0
End Sub

'--------------------------------------------------
' 卸载快捷键
'--------------------------------------------------
Public Sub UninstallShortcuts()
    On Error Resume Next
    Application.OnKey "^+L"
    Application.OnKey "^+Q"
    Application.OnKey "^+H"
    MsgBox "快捷键已卸载", vbInformation, "快捷键管理"
    On Error GoTo 0
End Sub

'==================================================
' 版本信息和帮助
'==================================================

'--------------------------------------------------
' 显示系统信息
'--------------------------------------------------
Public Sub ShowSystemInfo()
    Dim info As String
    
    info = "Excel智能布局优化系统 v3.1" & vbCrLf & vbCrLf
    info = info & "核心功能：" & vbCrLf
    info = info & "• 标题优先完整显示（新功能）" & vbCrLf
    info = info & "• 智能列宽优化" & vbCrLf
    info = info & "• 自动内容对齐" & vbCrLf
    info = info & "• 撤销/预览支持" & vbCrLf
    info = info & "• 自定义配置" & vbCrLf & vbCrLf
    
    info = info & "标题优先特性：" & vbCrLf
    info = info & "• 标题永不截断" & vbCrLf
    info = info & "• 自动换行调整" & vbCrLf
    info = info & "• 智能行高计算" & vbCrLf & vbCrLf
    
    info = info & "快捷键：" & vbCrLf
    info = info & "• Ctrl+Shift+L：启动优化" & vbCrLf
    info = info & "• Ctrl+Shift+Q：快速优化" & vbCrLf
    info = info & "• Ctrl+Z：撤销优化" & vbCrLf & vbCrLf
    
    info = info & "更新日期：2025年8月18日" & vbCrLf
    info = info & "作者：dadada"
    
    MsgBox info, vbInformation, "关于 Excel布局优化系统"
End Sub

'--------------------------------------------------
' 实现完整的GetUserConfiguration函数
'--------------------------------------------------
Private Function GetUserConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    Dim response As String
    Dim userChoice As VbMsgBoxResult
    
    ' 询问用户是否要自定义配置
    userChoice = MsgBox("是否要自定义布局配置？" & vbCrLf & vbCrLf & _
                       "选择'是'进行自定义设置" & vbCrLf & _
                       "选择'否'使用默认设置", _
                       vbYesNoCancel + vbQuestion, "Excel布局优化配置")
    
    If userChoice = vbCancel Then
        GetUserConfiguration = False
        Exit Function
    ElseIf userChoice = vbNo Then
        ' 使用默认配置
        GetUserConfiguration = True
        Exit Function
    End If
    
    ' 配置最大列宽
    response = InputBox("设置最大列宽（字符单位）" & vbCrLf & _
                       "范围: 30-100，当前: " & g_Config.MaxColumnWidth, _
                       "列宽配置", CStr(g_Config.MaxColumnWidth))
    
    If response <> "" And IsNumeric(response) Then
        Dim maxWidth As Double
        maxWidth = CDbl(response)
        If maxWidth >= 30 And maxWidth <= 100 Then
            g_Config.MaxColumnWidth = maxWidth
            g_Config.WrapThreshold = maxWidth
        End If
    End If
    
    ' 配置超长文本处理
    response = InputBox("设置超长文本列宽（字符单位）" & vbCrLf & _
                       "用于处理极长文本内容" & vbCrLf & _
                       "范围: 80-200，当前: " & g_Config.ExtremeTextWidth, _
                       "超长文本配置", CStr(g_Config.ExtremeTextWidth))
    
    If response <> "" And IsNumeric(response) Then
        Dim extremeWidth As Double
        extremeWidth = CDbl(response)
        If extremeWidth >= 80 And extremeWidth <= 200 Then
            g_Config.ExtremeTextWidth = extremeWidth
        End If
    End If
    
    ' 配置智能断行
    userChoice = MsgBox("是否启用智能断行？" & vbCrLf & vbCrLf & _
                       "启用后将在标点符号处优先换行，提升可读性", _
                       vbYesNo + vbQuestion, "智能断行配置")
    
    g_Config.SmartLineBreak = (userChoice = vbYes)
    
    ' 配置标题优先模式
    userChoice = MsgBox("是否启用标题优先模式？" & vbCrLf & vbCrLf & _
                       "启用后将确保标题完整显示，必要时自动换行", _
                       vbYesNo + vbQuestion, "标题优先配置")
    
    g_Config.HeaderPriority = (userChoice = vbYes)
    
    GetUserConfiguration = True
    Exit Function
    
ErrorHandler:
    GetUserConfiguration = False
End Function

'--------------------------------------------------
' 实现ShowPreviewDialog函数
'--------------------------------------------------
Private Function ShowPreviewDialog(info As PreviewInfo, targetRange As Range) As VbMsgBoxResult
    Dim message As String
    
    message = "布局优化预览" & vbCrLf & vbCrLf
    message = message & "优化区域: " & targetRange.Address & vbCrLf
    message = message & String(50, "-") & vbCrLf
    message = message & "• 总列数: " & info.TotalColumns & vbCrLf
    message = message & "• 需调整: " & info.ColumnsToAdjust & " 列" & vbCrLf
    
    If info.ColumnsNeedWrap > 0 Then
        message = message & "• 需换行: " & info.ColumnsNeedWrap & " 列"
        If g_Config.HeaderPriority Then
            message = message & "（包含标题）"
        End If
        message = message & vbCrLf
    End If
    
    message = message & "• 宽度范围: " & Format(info.MinWidth, "0.0") & _
              " - " & Format(info.MaxWidth, "0.0") & vbCrLf
    
    If g_Config.HeaderPriority Then
        message = message & "• 标题优先: 已启用" & vbCrLf
    End If
    
    If info.HasMergedCells Then
        message = message & "• 警告: 包含合并单元格（将跳过）" & vbCrLf
    End If
    
    If info.HasFormulas Then
        message = message & "• 提示: 包含公式" & vbCrLf
    End If
    
    message = message & String(50, "-") & vbCrLf
    message = message & "预计耗时: " & Format(info.EstimatedTime, "0.0") & " 秒" & vbCrLf & vbCrLf
    message = message & "是否继续？（处理中可按ESC中断）"
    
    ShowPreviewDialog = MsgBox(message, vbYesNoCancel + vbInformation, "Excel布局优化")
End Function

'--------------------------------------------------
' 实现SaveStateForUndo函数（完整版）
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
' 实现RestoreFromUndo函数
'--------------------------------------------------
Private Function RestoreFromUndo() As Boolean
    If Not g_HasUndoInfo Then
        RestoreFromUndo = False
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
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
    
    Application.ScreenUpdating = True
    RestoreFromUndo = True
    g_HasUndoInfo = False
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    RestoreFromUndo = False
End Function

'--------------------------------------------------
' 其他辅助函数
'--------------------------------------------------

Private Sub ResetCancelFlag()
    g_CancelOperation = False
End Sub

Private Function CheckForCancel() As Boolean
    ' 检查ESC键是否被按下
    DoEvents
    CheckForCancel = g_CancelOperation
End Function

Private Function HasMergedCells(targetRange As Range) As Boolean
    On Error Resume Next
    HasMergedCells = targetRange.MergeCells
    If IsNull(HasMergedCells) Then HasMergedCells = False
    On Error GoTo 0
End Function

Private Function IsHeaderRow(row1 As Range, row2 As Range) As Boolean
    ' 增强的表头检测逻辑 - 特别针对中文标题优化
    On Error GoTo ErrorHandler
    
    If row1 Is Nothing Then
        IsHeaderRow = False
        Exit Function
    End If
    
    ' 如果只有一行数据，假设第一行是标题
    If row2 Is Nothing Then
        IsHeaderRow = True
        Exit Function
    End If
    
    Dim score As Integer
    score = 0
    
    ' 检测1：第一行是否主要为文本，第二行是否主要为数字
    Dim textCount As Integer, numberCount As Integer
    Dim cell As Range
    
    For Each cell In row1.Cells
        If Not IsEmpty(cell.Value) Then
            If IsNumeric(cell.Value) Then
                numberCount = numberCount + 1
            Else
                textCount = textCount + 1
                ' 额外加分：包含中文字符
                Dim cellText As String
                cellText = CStr(cell.Value)
                Dim j As Integer
                For j = 1 To Len(cellText)
                    Dim charCode As Integer
                    charCode = Asc(Mid(cellText, j, 1))
                    If charCode > 127 Or charCode < 0 Then ' 中文字符
                        score = score + 1
                        Exit For
                    End If
                Next j
            End If
        End If
    Next cell
    
    If textCount > numberCount Then score = score + 2
    
    ' 检测2：第一行是否有格式化（加粗等）
    For Each cell In row1.Cells
        If cell.Font.Bold Or cell.Interior.ColorIndex <> xlNone Then
            score = score + 2
            Exit For
        End If
    Next cell
    
    ' 检测3：内容长度差异
    Dim avgLen1 As Double, avgLen2 As Double
    Dim count1 As Integer, count2 As Integer
    
    For Each cell In row1.Cells
        If Not IsEmpty(cell.Value) Then
            avgLen1 = avgLen1 + Len(CStr(cell.Value))
            count1 = count1 + 1
        End If
    Next cell
    
    For Each cell In row2.Cells
        If Not IsEmpty(cell.Value) Then
            avgLen2 = avgLen2 + Len(CStr(cell.Value))
            count2 = count2 + 1
        End If
    Next cell
    
    If count1 > 0 Then avgLen1 = avgLen1 / count1
    If count2 > 0 Then avgLen2 = avgLen2 / count2
    
    If avgLen1 > avgLen2 And avgLen1 > 3 Then score = score + 1
    
    ' 检测4：特殊关键词检测
    For Each cell In row1.Cells
        If Not IsEmpty(cell.Value) Then
            Dim testText As String
            testText = CStr(cell.Value)
            If InStr(testText, "数量") > 0 Or InStr(testText, "金额") > 0 Or _
               InStr(testText, "时间") > 0 Or InStr(testText, "日期") > 0 Or _
               InStr(testText, "名称") > 0 Or InStr(testText, "编号") > 0 Or _
               InStr(testText, "类型") > 0 Or InStr(testText, "状态") > 0 Then
                score = score + 2
                Exit For
            End If
        End If
    Next cell
    
    ' 检测5：位置因素 - 如果是第一行，增加权重
    If row1.Row = 1 Then
        score = score + 1
    End If
    
    ' 降低阈值，特别是对于包含中文的情况
    ' 得分>=2认为是表头（原来是2，现在保持不变但加分更容易）
    IsHeaderRow = (score >= 2)
    Exit Function
    
ErrorHandler:
    IsHeaderRow = True ' 出错时默认假设是标题
End Function

Private Sub ApplyColumnWidthOptimization(targetRange As Range, analyses() As ColumnAnalysisData)
    ' 优化列宽，但保持隐藏列的隐藏状态
    Dim i As Long
    For i = 1 To UBound(analyses)
        ' 只处理可见列，跳过隐藏列
        If Not targetRange.Columns(i).Hidden And Not analyses(i).HasMergedCells Then
            targetRange.Columns(i).ColumnWidth = analyses(i).OptimalWidth
        End If
    Next i
End Sub

Private Sub ApplyAlignmentOptimizationWithHeader(targetRange As Range, analyses() As ColumnAnalysisData, hasHeader As Boolean)
    ' 智能对齐优化，支持标题居中，但保护隐藏列
    On Error Resume Next
    
    Dim i As Long
    Dim col As Range
    
    For i = 1 To UBound(analyses)
        Set col = targetRange.Columns(i)
        
        ' 只处理可见列，跳过隐藏列
        If Not col.Hidden Then
            ' 如果有表头，设置标题行居中对齐
            If hasHeader Then
                col.Cells(1, 1).HorizontalAlignment = xlCenter
                col.Cells(1, 1).VerticalAlignment = xlCenter
                
                ' 为标题行设置轻微的格式增强（可选）
                If Not col.Cells(1, 1).Font.Bold Then
                    ' 只有在没有加粗的情况下才稍微增强格式
                    col.Cells(1, 1).Font.Size = col.Cells(1, 1).Font.Size + 0.5
                End If
                
                ' 如果标题较长，确保启用换行并保持居中
                If analyses(i).IsHeaderColumn And Len(analyses(i).HeaderText) > 8 Then
                    Dim headerWidth As Double
                    headerWidth = CalculateTextWidth(analyses(i).HeaderText, 12)
                    If headerWidth > col.ColumnWidth * 7 Then ' 如果文本宽度超过列宽
                        col.Cells(1, 1).WrapText = True
                    End If
                End If
            End If
            
            ' 根据数据类型设置数据行的对齐方式
            Dim startRow As Long
            startRow = IIf(hasHeader, 2, 1)
            
            Dim dataRange As Range
            Set dataRange = col.Cells(startRow, 1).Resize(targetRange.Rows.Count - startRow + 1, 1)
            
            ' 过滤出可见行进行处理
            Dim visibleDataRange As Range
            Set visibleDataRange = GetVisibleRange(dataRange)
            
            If Not visibleDataRange Is Nothing Then
                Select Case analyses(i).DataType
                    Case IntegerValue, DecimalValue, CurrencyValue, PercentageValue
                        ' 数值类型右对齐
                        visibleDataRange.HorizontalAlignment = xlRight
                        visibleDataRange.VerticalAlignment = xlCenter
                        
                    Case DateValue, TimeValue, DateTimeValue
                        ' 日期时间类型居中对齐
                        visibleDataRange.HorizontalAlignment = xlCenter
                        visibleDataRange.VerticalAlignment = xlCenter
                        
                    Case BooleanValue
                        ' 布尔值居中对齐
                        visibleDataRange.HorizontalAlignment = xlCenter
                        visibleDataRange.VerticalAlignment = xlCenter
                        
                    Case Else
                        ' 文本类型左对齐
                        visibleDataRange.HorizontalAlignment = xlLeft
                        visibleDataRange.VerticalAlignment = xlCenter
                End Select
            End If
            
            ' 如果列有标题且标题需要换行，确保标题居中
            If analyses(i).IsHeaderColumn And analyses(i).HeaderNeedWrap Then
                col.Cells(1, 1).WrapText = True
                col.Cells(1, 1).HorizontalAlignment = xlCenter
                col.Cells(1, 1).VerticalAlignment = xlCenter
            End If
        End If
    Next i
    
    On Error GoTo 0
End Sub

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
        If analyses(i).NeedWrap Then
            targetRange.Columns(i).WrapText = True
            
            ' 对超长文本列设置垂直对齐为顶部
            If analyses(i).IsHeaderColumn And Not IsEmpty(analyses(i).HeaderText) Then
                Dim headerCategory As TextLengthCategory
                headerCategory = ClassifyTextLength(analyses(i).HeaderText)
                
                If headerCategory >= TextLengthCategory.LongText Then
                    ' 超长文本设置顶部对齐提升可读性
                    targetRange.Columns(i).VerticalAlignment = xlTop
                End If
            End If
        End If
        
        ' 检查标题是否需要特殊处理
        If analyses(i).IsHeaderColumn Then
            ' 确保标题行启用换行（如果需要）
            If analyses(i).HeaderNeedWrap Then
                targetRange.Cells(1, i).WrapText = True
                hasHeaderAdjustment = True
                
                ' 计算需要的行高（增强版 - 支持超长文本）
                Dim calculatedHeight As Double
                If Not IsEmpty(analyses(i).HeaderText) Then
                    calculatedHeight = CalculateOptimalRowHeight(analyses(i).HeaderText, analyses(i).OptimalWidth)
                Else
                    calculatedHeight = analyses(i).HeaderRowHeight
                End If
                
                If calculatedHeight > maxHeaderHeight Then
                    maxHeaderHeight = calculatedHeight
                End If
            End If
        End If
    Next i
    
    ' 如果有标题需要调整，先设置标题行高（仅在标题行可见时）
    If hasHeaderAdjustment And Not targetRange.Rows(1).Hidden Then
        ' 限制最大行高避免界面问题
        If maxHeaderHeight > MAX_ROW_HEIGHT Then
            maxHeaderHeight = MAX_ROW_HEIGHT
        End If
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
        
        ' 只对可见行应用AutoFit
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

Private Function CalculateTextWidth(Text As String, fontSize As Single) As Double
    ' 精确的文本宽度计算 - 针对宋体12号字优化，特别是中文字符
    On Error GoTo ErrorHandler
    
    If Text = "" Then
        CalculateTextWidth = 0
        Exit Function
    End If
    
    Dim totalWidth As Double
    Dim i As Integer
    Dim char As String
    Dim charCode As Integer
    
    ' 宋体字符宽度系数（基于实际测量）
    Dim chineseCharWidth As Double
    Dim englishCharWidth As Double
    Dim numberCharWidth As Double
    
    ' 根据字体大小调整系数 - 特别针对12号宋体优化
    chineseCharWidth = (fontSize / 12) * 8.5  ' 中文字符在宋体下约8.5字符单位宽
    englishCharWidth = (fontSize / 12) * 4.8  ' 英文字符约4.8字符单位宽
    numberCharWidth = (fontSize / 12) * 5.2   ' 数字字符约5.2字符单位宽
    
    For i = 1 To Len(Text)
        char = Mid(Text, i, 1)
        charCode = Asc(char)
        
        If charCode > 127 Or charCode < 0 Then
            ' 中文字符
            totalWidth = totalWidth + chineseCharWidth
        ElseIf charCode >= 48 And charCode <= 57 Then
            ' 数字字符 (0-9)
            totalWidth = totalWidth + numberCharWidth
        Else
            ' 英文字符和符号
            totalWidth = totalWidth + englishCharWidth
        End If
    Next i
    
    CalculateTextWidth = totalWidth
    Exit Function
    
ErrorHandler:
    ' 简化计算作为备用
    CalculateTextWidth = Len(Text) * (fontSize / 12) * 6.5
End Function

Private Function AnalyzeHeaderWidth(headerText As String, maxWidth As Double) As Double
    On Error GoTo ErrorHandler
    
    If headerText = "" Then
        AnalyzeHeaderWidth = 0
        Exit Function
    End If
    
    ' 使用12号宋体计算标题的基本宽度（包含适当缓冲）
    Dim baseWidth As Double
    baseWidth = CalculateTextWidth(headerText, 12) + g_Config.TextBuffer
    
    ' 对于长标题增加额外缓冲，确保完整显示
    If Len(headerText) > 6 Then
        baseWidth = baseWidth + 2 ' 额外2字符单位缓冲
    End If
    
    ' 如果标题宽度在限制范围内，直接返回
    If baseWidth <= maxWidth Then
        AnalyzeHeaderWidth = baseWidth
    Else
        ' 标题太长，需要换行，返回最大宽度
        AnalyzeHeaderWidth = maxWidth
    End If
    
    Exit Function
    
ErrorHandler:
    AnalyzeHeaderWidth = g_Config.MinColumnWidth
End Function

Private Function CalculateHeaderRowHeight(headerText As String, columnWidth As Double) As Double
    On Error GoTo ErrorHandler
    
    If headerText = "" Then
        CalculateHeaderRowHeight = g_Config.HeaderMinHeight
        Exit Function
    End If
    
    ' 使用12号宋体计算需要的行数
    Dim textWidth As Double
    textWidth = CalculateTextWidth(headerText, 12)
    
    Dim linesNeeded As Long
    linesNeeded = Application.Max(1, Application.Ceiling(textWidth / columnWidth, 1))
    
    ' 限制最大行数
    If linesNeeded > g_Config.HeaderMaxWrapLines Then
        linesNeeded = g_Config.HeaderMaxWrapLines
    End If
    
    ' 计算行高（每行约15像素 + 间距）
    CalculateHeaderRowHeight = Application.Max(g_Config.HeaderMinHeight, linesNeeded * 18)
    
    Exit Function
    
ErrorHandler:
    CalculateHeaderRowHeight = g_Config.HeaderMinHeight
End Function

Private Function CalculateOptimalWidthWithHeader(analysis As ColumnAnalysisData) As WidthResult
    Dim result As WidthResult
    On Error GoTo ErrorHandler
    
    ' 如果不是标题列或没有启用标题优先，使用原有逻辑
    If Not analysis.IsHeaderColumn Or Not g_Config.HeaderPriority Then
        result = CalculateOptimalWidthEnhanced(analysis.MaxContentWidth, analysis.DataType)
        CalculateOptimalWidthWithHeader = result
        Exit Function
    End If
    
    ' 标题优先模式（增强版 - 支持超长文本）
    Dim headerRequiredWidth As Double
    Dim dataOptimalWidth As Double
    
    ' 检查标题是否为超长文本
    Dim headerCategory As TextLengthCategory
    headerCategory = ClassifyTextLength(analysis.HeaderText)
    
    ' 根据文本长度分类计算标题需要的宽度
    If headerCategory >= TextLengthCategory.LongText Then
        ' 长文本使用专门的计算函数
        headerRequiredWidth = CalculateExtremeTextWidth(analysis.HeaderText)
    Else
        ' 普通文本使用原有逻辑
        headerRequiredWidth = AnalyzeHeaderWidth(analysis.HeaderText, g_Config.MaxColumnWidth)
    End If
    
    ' 计算数据内容的最优宽度
    dataOptimalWidth = analysis.MaxContentWidth + g_Config.TextBuffer
    If dataOptimalWidth < g_Config.MinColumnWidth Then
        dataOptimalWidth = g_Config.MinColumnWidth
    End If
    
    ' 取两者中的较大值作为最终宽度
    result.FinalWidth = Application.Max(headerRequiredWidth, dataOptimalWidth)
    
    ' 检查是否需要换行（增强版）
    Dim headerTextWidth As Double
    headerTextWidth = CalculateTextWidth(analysis.HeaderText, 12)
    
    ' 超长文本的换行判断
    If headerCategory >= TextLengthCategory.VeryLongText Then
        ' 超长/极长文本强制换行
        result.NeedWrap = True
        result.FinalWidth = g_Config.ExtremeTextWidth
    ElseIf headerCategory = TextLengthCategory.LongText Then
        ' 长文本可选换行
        If headerTextWidth + g_Config.TextBuffer > result.FinalWidth Then
            result.NeedWrap = True
        Else
            result.NeedWrap = False
        End If
    Else
        ' 普通文本的换行判断
        If headerTextWidth + g_Config.TextBuffer > g_Config.MaxColumnWidth Then
            result.NeedWrap = True
            result.FinalWidth = g_Config.MaxColumnWidth
        Else
            result.NeedWrap = False
        End If
    End If
    
    ' 应用最终的边界控制
    If result.FinalWidth > g_Config.ExtremeTextWidth Then
        result.FinalWidth = g_Config.ExtremeTextWidth
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

Private Function SafeGetCellValue(cellValue As Variant) As String
    On Error Resume Next
    SafeGetCellValue = CStr(cellValue)
    On Error GoTo 0
End Function

Private Function CollectPreviewInfo(targetRange As Range) As PreviewInfo
    ' 收集预览信息（增强版 - 支持超长文本预览）
    Dim info As PreviewInfo
    info.AffectedCells = targetRange.Cells.Count
    info.TotalColumns = targetRange.Columns.Count
    info.HasMergedCells = HasMergedCells(targetRange)
    info.HasFormulas = HasFormulas(targetRange)
    
    ' 快速分析各列以收集预览信息
    Dim dataArray As Variant
    dataArray = SafeReadRangeToArray(targetRange)
    
    Dim col As Long
    Dim maxWidth As Double
    Dim minWidth As Double
    Dim headerCount As Long
    Dim extremeTextCount As Long
    
    maxWidth = 0
    minWidth = 999
    headerCount = 0
    extremeTextCount = 0
    
    For col = 1 To info.TotalColumns
        Dim analysis As ColumnAnalysisData
        analysis = AnalyzeColumnEnhanced(dataArray, col, targetRange.Rows.Count, targetRange.Columns(col))
        
        ' 统计调整信息
        If analysis.OptimalWidth <> targetRange.Columns(col).ColumnWidth Then
            info.ColumnsToAdjust = info.ColumnsToAdjust + 1
        End If
        
        If analysis.NeedWrap Or analysis.HeaderNeedWrap Then
            info.ColumnsNeedWrap = info.ColumnsNeedWrap + 1
        End If
        
        ' 记录宽度范围
        If analysis.OptimalWidth > maxWidth Then maxWidth = analysis.OptimalWidth
        If analysis.OptimalWidth < minWidth Then minWidth = analysis.OptimalWidth
        
        ' 统计标题列
        If analysis.IsHeaderColumn Then
            headerCount = headerCount + 1
            
            ' 检查是否包含超长文本
            If Not IsEmpty(analysis.HeaderText) Then
                Dim textCategory As TextLengthCategory
                textCategory = ClassifyTextLength(analysis.HeaderText)
                If textCategory >= TextLengthCategory.VeryLongText Then
                    extremeTextCount = extremeTextCount + 1
                End If
            End If
        End If
    Next col
    
    info.MaxWidth = maxWidth
    info.MinWidth = minWidth
    
    ' 估算处理时间（基于数据量和复杂度，包含超长文本处理时间）
    info.EstimatedTime = (info.AffectedCells / 10000) * 1.5
    If info.HasMergedCells Then info.EstimatedTime = info.EstimatedTime * 1.2
    If headerCount > 0 Then info.EstimatedTime = info.EstimatedTime * 1.1
    If extremeTextCount > 0 Then info.EstimatedTime = info.EstimatedTime * 1.3  ' 超长文本需要额外处理时间
    If info.EstimatedTime < 0.5 Then info.EstimatedTime = 0.5
    
    CollectPreviewInfo = info
End Function

Private Function HasFormulas(targetRange As Range) As Boolean
    On Error GoTo ErrorHandler
    
    Dim cell As Range
    For Each cell In targetRange
        If Not IsEmpty(cell.Value) And cell.HasFormula Then
            HasFormulas = True
            Exit Function
        End If
    Next cell
    
    HasFormulas = False
    Exit Function
    
ErrorHandler:
    HasFormulas = False
End Function

Private Function SafeReadRangeToArray(targetRange As Range) As Variant
    On Error GoTo ErrorHandler
    
    ' 参数验证
    If targetRange Is Nothing Then
        GoTo ErrorHandler
    End If
    
    ' 单个单元格特殊处理
    If targetRange.Cells.Count = 1 Then
        Dim singleArray(1 To 1, 1 To 1) As Variant
        singleArray(1, 1) = targetRange.Value2
        SafeReadRangeToArray = singleArray
        Exit Function
    End If
    
    ' 多单元格处理
    ' 使用Value2属性避免格式转换，提高性能
    SafeReadRangeToArray = targetRange.Value2
    
    ' 验证返回结果
    If Not IsArray(SafeReadRangeToArray) Then
        GoTo ErrorHandler
    End If
    
    Exit Function
    
ErrorHandler:
    ' 错误时返回安全的空数组
    Dim errorArray(1 To 1, 1 To 1) As Variant
    errorArray(1, 1) = ""
    SafeReadRangeToArray = errorArray
    
    ' 可选：记录错误信息用于调试
    Debug.Print "SafeReadRangeToArray 错误: " & Err.Description & " (错误号: " & Err.Number & ")"
End Function

' =================== 标题居中测试函数 ===================

Sub TestHeaderCentering()
    ' 测试标题居中功能的专用函数
    On Error GoTo ErrorHandler
    
    ' 创建测试数据
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 清空测试区域
    ws.Range("A1:D10").Clear
    
    ' 创建测试数据 - 中文标题
    ws.Range("A1").Value = "云仓新机数量"
    ws.Range("B1").Value = "云仓维修及数量"
    ws.Range("C1").Value = "产品类型"
    ws.Range("D1").Value = "处理状态"
    
    ' 添加一些数据行
    ws.Range("A2").Value = 100
    ws.Range("B2").Value = 50
    ws.Range("C2").Value = "手机"
    ws.Range("D2").Value = "已完成"
    
    ws.Range("A3").Value = 200
    ws.Range("B3").Value = 75
    ws.Range("C3").Value = "平板电脑"
    ws.Range("D3").Value = "处理中"
    
    ' 选择测试范围
    ws.Range("A1:D3").Select
    
    ' 应用优化
    ' Call OptimizeSelectedLayout ' Assuming this calls OptimizeLayout or a similar function
    Call OptimizeLayout
    
    ' 检查结果
    Dim headerRange As Range
    Set headerRange = ws.Range("A1:D1")
    
    Dim resultMsg As String
    resultMsg = "标题居中测试结果:" & vbCrLf & vbCrLf
    
    Dim cell As Range
    For Each cell In headerRange
        resultMsg = resultMsg & "'" & cell.Value & "' - 对齐方式: "
        Select Case cell.HorizontalAlignment
            Case xlCenter
                resultMsg = resultMsg & "居中 ✓"
            Case xlLeft
                resultMsg = resultMsg & "左对齐 ✗"
            Case xlRight
                resultMsg = resultMsg & "右对齐 ✗"
            Case Else
                resultMsg = resultMsg & "其他(" & cell.HorizontalAlignment & ") ✗"
        End Select
        resultMsg = resultMsg & vbCrLf
    Next cell
    
    resultMsg = resultMsg & vbCrLf & "测试完成！请检查 A1:D3 区域的标题是否已居中显示。"
    MsgBox resultMsg, vbInformation, "标题居中测试"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "测试过程中发生错误: " & Err.Description, vbCritical, "错误"
End Sub

Sub TestLongHeaderDisplay()
    ' 测试长标题完整显示功能的专用函数
    On Error GoTo ErrorHandler
    
    ' 创建测试数据
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 清空测试区域
    ws.Range("A1:E10").Clear
    
    ' 创建测试数据 - 用户提到的具体长标题
    ws.Range("A1").Value = "云仓新机数量"
    ws.Range("B1").Value = "公司新机数量"
    ws.Range("C1").Value = "公司维修机数量"
    ws.Range("D1").Value = "发货日期（退货签收）时间"
    ws.Range("E1").Value = "产品处理状态"
    
    ' 添加一些数据行
    ws.Range("A2").Value = 150
    ws.Range("B2").Value = 80
    ws.Range("C2").Value = 25
    ws.Range("D2").Value = "2025-08-18"
    ws.Range("E2").Value = "已处理"
    
    ws.Range("A3").Value = 200
    ws.Range("B3").Value = 120
    ws.Range("C3").Value = 40
    ws.Range("D3").Value = "2025-08-19"
    ws.Range("E3").Value = "处理中"
    
    ' 应用优化前记录列宽
    Dim originalWidths(1 To 5) As Double
    Dim i As Integer
    For i = 1 To 5
        originalWidths(i) = ws.Columns(i).ColumnWidth
    Next i
    
    ' 选择测试范围
    ws.Range("A1:E3").Select
    
    ' 应用优化
    ' Call OptimizeSelectedLayout ' Assuming this calls OptimizeLayout or a similar function
    Call OptimizeLayout
    
    ' 检查结果并生成报告
    Dim resultMsg As String
    resultMsg = "长标题显示测试结果:" & vbCrLf & vbCrLf
    
    Dim headers As Variant
    headers = Array("云仓新机数量", "公司新机数量", "公司维修机数量", "发货日期（退货签收）时间", "产品处理状态")
    
    For i = 1 To 5
        Dim currentWidth As Double
        currentWidth = ws.Columns(i).ColumnWidth
        
        resultMsg = resultMsg & "列" & i & ": " & headers(i - 1) & vbCrLf
        resultMsg = resultMsg & "  原宽度: " & Format(originalWidths(i), "0.0") & " → 新宽度: " & Format(currentWidth, "0.0")
        
        ' 检查是否有换行
        If ws.Cells(1, i).WrapText Then
            resultMsg = resultMsg & " (已启用换行)"
        End If
        
        ' 检查对齐方式
        Select Case ws.Cells(1, i).HorizontalAlignment
            Case xlCenter
                resultMsg = resultMsg & " [居中]"
            Case xlLeft
                resultMsg = resultMsg & " [左对齐]"
            Case xlRight
                resultMsg = resultMsg & " [右对齐]"
        End Select
        
        resultMsg = resultMsg & vbCrLf
    Next i
    
    resultMsg = resultMsg & vbCrLf & "请检查 A1:E3 区域的标题是否完整显示！"
    MsgBox resultMsg, vbInformation, "长标题显示测试"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "测试过程中发生错误: " & Err.Description, vbCritical, "错误"
End Sub

' =================== 辅助函数：保护隐藏行列 ===================

Private Function GetVisibleRange(inputRange As Range) As Range
    ' 从指定范围中提取可见的单元格
    On Error GoTo ErrorHandler
    
    Dim visibleCells As Range
    Set visibleCells = inputRange.SpecialCells(xlCellTypeVisible)
    Set GetVisibleRange = visibleCells
    
    Exit Function

ErrorHandler:
    ' 如果没有可见单元格或发生错误，返回Nothing
    Set GetVisibleRange = Nothing
End Function

Sub TestHiddenCellsProtection()
    ' 测试隐藏行列保护功能
    On Error GoTo ErrorHandler
    
    ' 创建测试数据
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 清空测试区域
    ws.Range("A1:E10").Clear
    ws.Range("A1:E10").RowHeight = -1 ' 重置行高为自动
    ws.Columns("A:E").ColumnWidth = 8.43 ' 重置列宽
    ws.Columns("A:E").Hidden = False ' 显示所有列
    ws.Rows("1:10").Hidden = False ' 显示所有行
    
    ' 创建测试数据
    ws.Range("A1").Value = "列A标题"
    ws.Range("B1").Value = "列B标题"
    ws.Range("C1").Value = "列C标题"
    ws.Range("D1").Value = "列D标题"
    ws.Range("E1").Value = "列E标题"
    
    ' 添加数据行
    ws.Range("A2").Value = "数据A2"
    ws.Range("B2").Value = "数据B2"
    ws.Range("C2").Value = "数据C2"
    ws.Range("D2").Value = "数据D2"
    ws.Range("E2").Value = "数据E2"
    
    ws.Range("A3").Value = "数据A3"
    ws.Range("B3").Value = "数据B3"
    ws.Range("C3").Value = "数据C3"
    ws.Range("D3").Value = "数据D3"
    ws.Range("E3").Value = "数据E3"
    
    ' 隐藏列C和行3
    ws.Columns("C").Hidden = True
    ws.Rows("3").Hidden = True
    
    ' 记录优化前的状态
    Dim beforeOptimization As String
    beforeOptimization = "优化前状态:" & vbCrLf
    beforeOptimization = beforeOptimization & "列C隐藏: " & ws.Columns("C").Hidden & vbCrLf
    beforeOptimization = beforeOptimization & "行3隐藏: " & ws.Rows("3").Hidden & vbCrLf
    beforeOptimization = beforeOptimization & "列A宽度: " & Format(ws.Columns("A").ColumnWidth, "0.0") & vbCrLf
    beforeOptimization = beforeOptimization & "列C宽度: " & Format(ws.Columns("C").ColumnWidth, "0.0") & vbCrLf
    
    ' 选择范围并应用优化
    ws.Range("A1:E3").Select
    Call OptimizeLayout
    
    ' 检查优化后的状态
    Dim afterOptimization As String
    afterOptimization = vbCrLf & "优化后状态:" & vbCrLf
    afterOptimization = afterOptimization & "列C隐藏: " & ws.Columns("C").Hidden & vbCrLf
    afterOptimization = afterOptimization & "行3隐藏: " & ws.Rows("3").Hidden & vbCrLf
    afterOptimization = afterOptimization & "列A宽度: " & Format(ws.Columns("A").ColumnWidth, "0.0") & vbCrLf
    afterOptimization = afterOptimization & "列C宽度: " & Format(ws.Columns("C").ColumnWidth, "0.0") & vbCrLf
    
    ' 验证结果
    Dim testResult As String
    testResult = vbCrLf & "测试结果:" & vbCrLf
    
    If ws.Columns("C").Hidden Then
        testResult = testResult & "✓ 列C保持隐藏状态" & vbCrLf
    Else
        testResult = testResult & "✗ 列C隐藏状态被取消" & vbCrLf
    End If
    
    If ws.Rows("3").Hidden Then
        testResult = testResult & "✓ 行3保持隐藏状态" & vbCrLf
    Else
        testResult = testResult & "✗ 行3隐藏状态被取消" & vbCrLf
    End If
    
    ' 显示完整测试报告
    Dim fullReport As String
    fullReport = "隐藏行列保护测试" & vbCrLf
    fullReport = fullReport & "===================" & vbCrLf
    fullReport = fullReport & beforeOptimization
    fullReport = fullReport & afterOptimization
    fullReport = fullReport & testResult
    
    MsgBox fullReport, vbInformation, "隐藏行列保护测试结果"

    Exit Sub
    
ErrorHandler:
    MsgBox "测试过程中发生错误: " & Err.Description, vbCritical, "错误"
End Sub

' =================== 超长文本处理函数（新增） ===================

'--------------------------------------------------
' 分类文本长度
'--------------------------------------------------
Private Function ClassifyTextLength(text As String) As TextLengthCategory
    Dim textLength As Long
    textLength = Len(text)
    
    If textLength <= 20 Then
        ClassifyTextLength = TextLengthCategory.ShortText
    ElseIf textLength <= 50 Then
        ClassifyTextLength = TextLengthCategory.MediumText
    ElseIf textLength <= 100 Then
        ClassifyTextLength = TextLengthCategory.LongText
    ElseIf textLength <= 200 Then
        ClassifyTextLength = TextLengthCategory.VeryLongText
    Else
        ClassifyTextLength = TextLengthCategory.ExtremeText
    End If
End Function

'--------------------------------------------------
' 查找智能断行点
'--------------------------------------------------
Private Function FindBreakPoints(text As String) As Collection
    Dim breaks As New Collection
    Dim i As Long
    Dim textLength As Long
    textLength = Len(text)
    
    ' 优先在标点符号处断行
    Dim punctuation As String
    punctuation = "，。；：！？,;:!?"
    
    For i = 1 To textLength
        Dim char As String
        char = Mid(text, i, 1)
        
        ' 检查是否为标点符号
        If InStr(punctuation, char) > 0 Then
            breaks.Add i
        ' 其次在空格处断行
        ElseIf char = " " Then
            breaks.Add i
        End If
    Next i
    
    Set FindBreakPoints = breaks
End Function

'--------------------------------------------------
' 计算智能换行布局
'--------------------------------------------------
Private Function CalculateWrapLayout(text As String, maxWidth As Double) As WrapLayout
    Dim layout As WrapLayout
    
    On Error GoTo ErrorHandler
    
    ' 获取文本总宽度
    Dim totalWidth As Double
    totalWidth = CalculateTextWidth(text, 11)
    
    ' 如果不需要换行
    If totalWidth <= maxWidth Then
        layout.TotalLines = 1
        layout.OptimalRowHeight = MIN_ROW_HEIGHT
        layout.NeedWrap = False
        CalculateWrapLayout = layout
        Exit Function
    End If
    
    ' 需要换行的情况
    layout.NeedWrap = True
    
    ' 估算需要的行数
    Dim estimatedLines As Long
    estimatedLines = Application.Ceiling(totalWidth / maxWidth, 1)
    
    ' 限制最大行数
    If estimatedLines > g_Config.MaxWrapLines Then
        estimatedLines = g_Config.MaxWrapLines
    End If
    
    layout.TotalLines = estimatedLines
    
    ' 计算行高（每行约18像素包含间距）
    layout.OptimalRowHeight = Application.Max(MIN_ROW_HEIGHT, estimatedLines * 18)
    
    ' 限制最大行高
    If layout.OptimalRowHeight > MAX_ROW_HEIGHT Then
        layout.OptimalRowHeight = MAX_ROW_HEIGHT
    End If
    
    CalculateWrapLayout = layout
    Exit Function
    
ErrorHandler:
    ' 错误情况下返回安全默认值
    layout.TotalLines = 1
    layout.OptimalRowHeight = MIN_ROW_HEIGHT
    layout.NeedWrap = False
    CalculateWrapLayout = layout
End Function

'--------------------------------------------------
' 计算超长文本的最优行高
'--------------------------------------------------
Private Function CalculateOptimalRowHeight(text As String, columnWidth As Double) As Double
    On Error GoTo ErrorHandler
    
    Dim baseHeight As Double
    baseHeight = MIN_ROW_HEIGHT
    
    ' 计算文本需要的行数
    Dim textWidth As Double
    textWidth = CalculateTextWidth(text, 11)
    
    Dim linesNeeded As Long
    linesNeeded = Application.Max(1, Application.Ceiling(textWidth / columnWidth, 1))
    
    ' 限制最大行数
    If linesNeeded > g_Config.MaxWrapLines Then
        linesNeeded = g_Config.MaxWrapLines
    End If
    
    ' 计算总行高（每行18像素 + 1.2倍行距）
    Dim totalHeight As Double
    totalHeight = baseHeight + (linesNeeded - 1) * 18 * 1.2
    
    ' 应用行高限制
    If totalHeight > MAX_ROW_HEIGHT Then
        totalHeight = MAX_ROW_HEIGHT
    End If
    
    CalculateOptimalRowHeight = totalHeight
    Exit Function
    
ErrorHandler:
    CalculateOptimalRowHeight = MIN_ROW_HEIGHT
End Function

'--------------------------------------------------
' 超长文本列宽计算（带分级处理）
'--------------------------------------------------
Private Function CalculateExtremeTextWidth(text As String) As Double
    On Error GoTo ErrorHandler
    
    Dim textCategory As TextLengthCategory
    textCategory = ClassifyTextLength(text)
    
    Select Case textCategory
        Case TextLengthCategory.ShortText
            ' 短文本：内容宽度+缓冲
            CalculateExtremeTextWidth = CalculateTextWidth(text, 11) + g_Config.TextBuffer
            
        Case TextLengthCategory.MediumText
            ' 中等文本：内容宽度+缓冲（上限70）
            Dim mediumWidth As Double
            mediumWidth = CalculateTextWidth(text, 11) + g_Config.TextBuffer
            CalculateExtremeTextWidth = Application.Min(mediumWidth, 70)
            
        Case TextLengthCategory.LongText
            ' 长文本：扩展至100
            CalculateExtremeTextWidth = 100
            
        Case TextLengthCategory.VeryLongText
            ' 超长文本：固定120
            CalculateExtremeTextWidth = g_Config.ExtremeTextWidth
            
        Case TextLengthCategory.ExtremeText
            ' 极长文本：固定120
            CalculateExtremeTextWidth = g_Config.ExtremeTextWidth
            
        Case Else
            ' 默认情况
            CalculateExtremeTextWidth = g_Config.MaxColumnWidth
    End Select
    
    ' 应用最小宽度限制
    If CalculateExtremeTextWidth < g_Config.MinColumnWidth Then
        CalculateExtremeTextWidth = g_Config.MinColumnWidth
    End If
    
    Exit Function
    
ErrorHandler:
    CalculateExtremeTextWidth = g_Config.MaxColumnWidth
End Function

' =================== 测试超长文本处理功能 ===================

Sub TestExtremeTextHandling()
    ' 测试超长文本处理功能
    On Error GoTo ErrorHandler
    
    ' 创建测试数据
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 清空测试区域
    ws.Range("A1:E10").Clear
    
    ' 创建不同长度的测试文本
    ws.Range("A1").Value = "短标题"  ' 短文本
    ws.Range("B1").Value = "这是一个中等长度的标题文本示例"  ' 中等文本
    ws.Range("C1").Value = "这是一个比较长的标题文本示例，用来测试长文本的处理效果，看看是否能够正确识别和处理长文本内容，确保显示效果良好"  ' 长文本
    ws.Range("D1").Value = "这是一个超长的标题文本示例，专门用来测试系统对于超长文本的处理能力，包括智能换行、行高调整等功能，确保即使是很长的文本内容也能够在表格中正确显示，不会出现截断或者格式混乱的问题，同时保持良好的可读性和美观性"  ' 超长文本
    ws.Range("E1").Value = "这是一个极长的标题文本示例，专门设计用来测试系统在处理极端长度文本时的表现，包括但不限于：智能换行处理、行高自动调整、列宽优化计算、文本截断保护、格式保持、可读性优化、性能控制等多个方面的功能，确保系统能够在各种极端情况下都能够稳定运行并提供良好的用户体验，同时不会因为文本过长而导致程序崩溃或者性能问题，这是一个综合性的测试案例"  ' 极长文本
    
    ' 添加一些数据行
    ws.Range("A2").Value = 100
    ws.Range("B2").Value = 200
    ws.Range("C2").Value = 300
    ws.Range("D2").Value = 400
    ws.Range("E2").Value = 500
    
    ' 记录优化前的列宽
    Dim originalWidths(1 To 5) As Double
    Dim i As Integer
    For i = 1 To 5
        originalWidths(i) = ws.Columns(i).ColumnWidth
    Next i
    
    ' 选择测试范围并应用优化
    ws.Range("A1:E2").Select
    Call OptimizeLayout
    
    ' 生成测试报告
    Dim resultMsg As String
    resultMsg = "超长文本处理测试结果:" & vbCrLf & vbCrLf
    
    Dim headers As Variant
    headers = Array("短标题", "中等长度标题", "长文本标题", "超长文本标题", "极长文本标题")
    
    For i = 1 To 5
        Dim currentWidth As Double
        currentWidth = ws.Columns(i).ColumnWidth
        
        Dim textLength As Long
        textLength = Len(ws.Cells(1, i).Value)
        
        resultMsg = resultMsg & "列" & i & " (长度:" & textLength & "字符):" & vbCrLf
        resultMsg = resultMsg & "  原宽度: " & Format(originalWidths(i), "0.0")
        resultMsg = resultMsg & " → 新宽度: " & Format(currentWidth, "0.0")
        
        ' 检查是否有换行
        If ws.Cells(1, i).WrapText Then
            resultMsg = resultMsg & " [换行]"
        End If
        
        ' 检查行高
        Dim rowHeight As Double
        rowHeight = ws.Rows(1).RowHeight
        resultMsg = resultMsg & " 行高:" & Format(rowHeight, "0.0")
        
        resultMsg = resultMsg & vbCrLf
    Next i
    
    resultMsg = resultMsg & vbCrLf & "请检查 A1:E2 区域的文本是否完整显示！"
    MsgBox resultMsg, vbInformation, "超长文本处理测试结果"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "测试过程中发生错误: " & Err.Description, vbCritical, "错误"
End Sub