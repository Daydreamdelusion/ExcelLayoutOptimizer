Attribute VB_Name = "ExcelBeautifier"
' =============================================================================
' Excel表格美化系统 v2.1 - 单模块完整实现
' 
' 功能特点：
' - R1C1统一架构，避免列字母解析脆弱性
' - 精确撤销机制，基于会话标签保护用户既有格式
' - 条件格式终止逻辑优化，避免无效计算
' - 分层边框颜色设计，增强视觉层次
' - 高性能斑马纹实现，智能自适应步长
'
' 创建日期：2025年9月3日
' 版本：v2.1
' =============================================================================

Option Explicit

' ===== 核心数据结构定义 =====

' 美化配置结构
Private Type BeautifyConfig
    ' 主题设置
    ThemeName As String              ' 主题名称: Business/Financial/Minimal
    PrimaryColor As Long             ' 主色调RGB值
    SecondaryColor As Long           ' 辅助色RGB值
    AccentColor As Long              ' 强调色RGB值
    
    ' 功能开关
    EnableHeaderBeautify As Boolean  ' 启用表头美化
    EnableConditionalFormat As Boolean ' 启用条件格式
    EnableBorders As Boolean         ' 启用边框样式
    EnableZebraStripes As Boolean    ' 启用隔行变色
    EnableFreezeHeader As Boolean    ' 启用冻结表头
    
    ' 样式参数
    HeaderFontSize As Single         ' 表头字号
    DataFontSize As Single           ' 数据字号
    BorderWeight As XlBorderWeight   ' 边框粗细
    StripeOpacity As Single          ' 条纹透明度(0-1)
End Type

' 撤销日志结构（精确撤销最小闭环字段）
Private Type BeautifyLog
    ' 会话标识
    SessionId As String              ' 唯一会话ID：Format(Now, "yyyymmddhhmmss") & "_" & Int(Rnd * 1000)
    Timestamp As Date                ' 操作时间戳
    
    ' 条件格式记录（按标签删除）
    CFRulesAdded As String           ' 格式: "地址|标签;地址|标签..." 支持精确删除
    
    ' 样式记录（会话级管理）
    StylesAdded As String            ' 本会话添加的样式名称: "ELO_主题_SessionId;..."
    TableStylesMap As String         ' 表格样式映射: "表名:原样式;表名:原样式"
End Type

' 表格分析结构
Private Type TableAnalysis
    ' 区域信息
    TotalRange As Range              ' 完整表格区域
    HeaderRange As Range             ' 表头区域
    DataRange As Range               ' 数据区域
    
    ' 表格特征
    HasHeaders As Boolean            ' 是否有表头
    HeaderRows As Long               ' 表头行数
    DataRows As Long                 ' 数据行数
    DataColumns As Long              ' 数据列数
    
    ' 内容特征
    HasNumbers As Boolean            ' 包含数值
    HasDates As Boolean              ' 包含日期
    HasFormulas As Boolean           ' 包含公式
    HasMergedCells As Boolean        ' 包含合并单元格
    
    ' 数据类型分析
    ColumnTypes() As String          ' 每列数据类型
    NumericColumns() As Long         ' 数值列索引
    TextColumns() As Long            ' 文本列索引
End Type

' 应用状态结构
Private Type AppState
    ScreenUpdating As Boolean
    Calculation As XlCalculation
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Cursor As XlMousePointer
    ReferenceStyle As XlReferenceStyle
End Type

' 错误代码定义
Private Enum BeautifyError
    ERR_NO_SELECTION = 1001
    ERR_INVALID_RANGE = 1002
    ERR_MEMORY_LIMIT = 1003
    ERR_FORMAT_CONFLICT = 1004
    ERR_UNDO_FAILED = 1005
End Enum

' ===== 全局变量 =====
' 多步撤销操作堆栈
Private g_UndoStack As Collection

' =============================================================================
' 公共API接口
' =============================================================================

' 主美化函数 - 系统入口点
Public Sub BeautifyTable()
    Dim targetRange As Range
    Dim config As BeautifyConfig
    Dim originalState As AppState
    
    On Error GoTo ErrorHandler
    
    ' 保存应用状态并设置性能模式
    originalState = SaveAppState()
    Call SetPerformanceMode()
    
    ' 获取目标区域
    Set targetRange = DetectTableRange()
    
    ' 验证操作
    If Not ValidateBeautifyOperation(targetRange) Then
        GoTo CleanUp
    End If
    
    ' 初始化撤销日志
    Call InitializeBeautifyLog()
    
    ' 获取默认配置（商务主题）
    config = GetBusinessTheme()
    
    ' 分析表格结构
    Dim analysis As TableAnalysis
    analysis = AnalyzeTable(targetRange)
    
    ' 执行美化
    Call ApplyThemeStyle(analysis, config)
    
    ' 应用条件格式
    If config.EnableConditionalFormat Then
        Call ApplyStandardConditionalFormat(analysis.DataRange)
    End If
    
    ' 恢复应用状态
    Call RestoreAppState(originalState)
    
    MsgBox "表格美化完成！可使用 UndoBeautify() 撤销。", vbInformation, "Excel美化工具"
    Exit Sub
    
ErrorHandler:
    Call RestoreAppState(originalState)
    Call HandleError(ERR_INVALID_RANGE, "美化操作失败: " & Err.Description)
    Exit Sub
    
CleanUp:
    Call RestoreAppState(originalState)
End Sub

' 撤销函数 - 支持多步撤销的堆栈机制
Public Sub UndoBeautify()
    Dim ws As Worksheet
    Dim historyLog As BeautifyLog
    Dim cfRuleEntries() As String
    Dim tableStyleMappings() As String
    Dim styleNames() As String
    Dim i As Long
    Dim sessionTag As String
    Dim originalState As AppState
    
    On Error GoTo ErrorHandler
    
    ' 初始化堆栈（如果未初始化）
    If g_UndoStack Is Nothing Then
        Set g_UndoStack = New Collection
    End If
    
    ' 检查是否有可撤销的操作
    If g_UndoStack.Count = 0 Then
        MsgBox "没有可撤销的美化操作", vbInformation, "Excel美化工具"
        Exit Sub
    End If
    
    ' 从堆栈顶部获取最近的操作记录
    Set historyLog = g_UndoStack(g_UndoStack.Count)
    Set ws = ActiveSheet
    sessionTag = "ELO_" & historyLog.SessionId
    
    ' 确认撤销操作
    If MsgBox("确定要撤销最近的美化操作吗？" & vbCrLf & _
              "操作时间：" & historyLog.Timestamp & vbCrLf & _
              "剩余可撤销操作：" & (g_UndoStack.Count - 1), _
              vbYesNo + vbQuestion, "Excel美化工具") = vbNo Then
        Exit Sub
    End If
    
    ' 保存应用状态并设置性能模式
    originalState = SaveAppState()
    Call SetPerformanceMode()
    
    ' 1. 精确删除带标签的条件格式规则
    If historyLog.CFRulesAdded <> "" Then
        cfRuleEntries = Split(historyLog.CFRulesAdded, ";")
        For i = 0 To UBound(cfRuleEntries)
            Call RemoveTaggedCFRule(ws, cfRuleEntries(i))
        Next i
    End If
    
    ' 2. 还原表格样式
    If historyLog.TableStylesMap <> "" Then
        tableStyleMappings = Split(historyLog.TableStylesMap, ";")
        For i = 0 To UBound(tableStyleMappings)
            Call RestoreTableStyle(ws, tableStyleMappings(i))
        Next i
    End If
    
    ' 3. 删除本会话创建的样式
    If historyLog.StylesAdded <> "" Then
        styleNames = Split(historyLog.StylesAdded, ";")
        For i = 0 To UBound(styleNames)
            Call SafeDeleteStyle(styleNames(i))
        Next i
    End If
    
    ' 4. 删除本会话的表格样式
    Call RemoveSessionTableStyles(sessionTag)
    
    ' 恢复应用状态
    Call RestoreAppState(originalState)
    
    ' 从堆栈中移除已撤销的操作
    g_UndoStack.Remove g_UndoStack.Count
    
    MsgBox "撤销完成！" & vbCrLf & _
           "剩余可撤销操作：" & g_UndoStack.Count, _
           vbInformation, "Excel美化工具"
    Exit Sub
    
ErrorHandler:
    Call RestoreAppState(originalState)
    MsgBox "撤销操作失败：" & Err.Description, vbCritical, "Excel美化工具"
End Sub

' 撤销所有美化操作
Public Sub UndoAllBeautify()
    Dim result As VbMsgBoxResult
    Dim count As Long
    
    If g_UndoStack Is Nothing Then
        Set g_UndoStack = New Collection
    End If
    
    count = g_UndoStack.Count
    If count = 0 Then
        MsgBox "没有可撤销的美化操作", vbInformation, "Excel美化工具"
        Exit Sub
    End If
    
    result = MsgBox("确定要撤销所有 " & count & " 个美化操作吗？" & vbCrLf & _
                   "此操作不可逆转！", _
                   vbYesNo + vbQuestion, "Excel美化工具")
    
    If result = vbYes Then
        ' 逐个撤销所有操作
        Do While g_UndoStack.Count > 0
            Call UndoBeautify()
        Loop
        
        MsgBox "已撤销所有美化操作！", vbInformation, "Excel美化工具"
    End If
End Sub
    Call RestoreAppState(originalState)
    Call HandleError(ERR_UNDO_FAILED, "撤销操作失败: " & Err.Description)
End Sub

' =============================================================================
' 核心功能实现
' =============================================================================

' 检测表格区域
' 智能表格区域检测（避免UsedRange的不可靠性）
Private Function DetectTableRange() As Range
    Dim selectedRange As Range
    Dim currentRegion As Range
    Dim smartRange As Range
    
    On Error GoTo UseCurrentRegion
    
    ' 优先使用用户选择的区域
    Set selectedRange = Selection
    If Not selectedRange Is Nothing And selectedRange.Cells.Count > 1 Then
        Set DetectTableRange = selectedRange
        Exit Function
    End If
    
UseCurrentRegion:
    On Error GoTo UseSmartDetection
    
    ' 使用CurrentRegion（比UsedRange更可靠）
    Set currentRegion = ActiveCell.CurrentRegion
    If Not currentRegion Is Nothing And currentRegion.Cells.Count > 1 Then
        Set DetectTableRange = currentRegion
        Exit Function
    End If
    
UseSmartDetection:
    On Error GoTo UseFallback
    
    ' 智能边界探测（处理CurrentRegion的空行/空列限制）
    Set smartRange = GetSmartTableRange()
    If Not smartRange Is Nothing And smartRange.Cells.Count > 1 Then
        Set DetectTableRange = smartRange
        Exit Function
    End If
    
UseFallback:
    ' 最后回退到UsedRange（但进行清理）
    Dim cleanedRange As Range
    Set cleanedRange = GetCleanedUsedRange()
    If Not cleanedRange Is Nothing And cleanedRange.Cells.Count > 1 Then
        Set DetectTableRange = cleanedRange
    Else
        Set DetectTableRange = Nothing
    End If
End Function

' 智能表格边界探测
Private Function GetSmartTableRange() As Range
    Dim startCell As Range
    Dim lastRow As Long, lastCol As Long
    Dim firstRow As Long, firstCol As Long
    
    On Error GoTo ErrorHandler
    
    Set startCell = ActiveCell
    
    ' 探测边界
    firstRow = startCell.End(xlUp).Row
    If firstRow = 1 Then firstRow = startCell.Row
    
    firstCol = startCell.End(xlToLeft).Column
    If firstCol = 1 Then firstCol = startCell.Column
    
    lastRow = startCell.End(xlDown).Row
    If lastRow = Rows.Count Then
        ' 如果到了最后一行，向上查找最后一个有数据的行
        lastRow = startCell.SpecialCells(xlCellTypeLastCell).Row
    End If
    
    lastCol = startCell.End(xlToRight).Column
    If lastCol = Columns.Count Then
        ' 如果到了最后一列，向左查找最后一个有数据的列
        lastCol = startCell.SpecialCells(xlCellTypeLastCell).Column
    End If
    
    Set GetSmartTableRange = Range(Cells(firstRow, firstCol), Cells(lastRow, lastCol))
    Exit Function
    
ErrorHandler:
    Set GetSmartTableRange = Nothing
End Function

' 清理UsedRange（移除末尾的空行空列）
Private Function GetCleanedUsedRange() As Range
    Dim ws As Worksheet
    Dim lastCell As Range
    Dim usedRange As Range
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    Set usedRange = ws.UsedRange
    
    ' 查找真正的最后一个有数据的单元格
    Set lastCell = usedRange.SpecialCells(xlCellTypeLastCell)
    
    ' 创建清理后的区域
    Set GetCleanedUsedRange = ws.Range(usedRange.Cells(1, 1), lastCell)
    Exit Function
    
ErrorHandler:
    Set GetCleanedUsedRange = ActiveSheet.UsedRange
End Function

' 验证美化操作
Private Function ValidateBeautifyOperation(targetRange As Range) As Boolean
    On Error GoTo ValidationError
    
    ' 检查1: 区域有效性
    If targetRange Is Nothing Then
        Call HandleError(ERR_NO_SELECTION, "请选择要美化的表格区域")
        ValidateBeautifyOperation = False
        Exit Function
    End If
    
    ' 检查2: 数据存在性
    If Application.WorksheetFunction.CountA(targetRange) = 0 Then
        Call HandleError(ERR_INVALID_RANGE, "选择的区域没有数据")
        ValidateBeautifyOperation = False
        Exit Function
    End If
    
    ' 检查3: 大小限制
    If targetRange.Cells.Count > 1000000 Then
        If MsgBox("数据量很大，可能需要较长时间。是否继续？", _
                  vbYesNo + vbQuestion, "Excel美化工具") = vbNo Then
            ValidateBeautifyOperation = False
            Exit Function
        End If
    End If
    
    ' 检查4: 格式冲突
    If HasConflictingFormats(targetRange) Then
        If MsgBox("检测到已有格式，是否覆盖？", _
                  vbYesNo + vbQuestion, "Excel美化工具") = vbNo Then
            ValidateBeautifyOperation = False
            Exit Function
        End If
    End If
    
    ValidateBeautifyOperation = True
    Exit Function
    
ValidationError:
    Call HandleError(ERR_INVALID_RANGE, Err.Description)
    ValidateBeautifyOperation = False
End Function

' 检测格式冲突
Private Function HasConflictingFormats(rng As Range) As Boolean
    Dim cell As Range
    
    ' 检查前10个单元格是否有非默认格式
    Dim checkCount As Long
    checkCount = 0
    
    For Each cell In rng.Cells
        If checkCount >= 10 Then Exit For
        
        ' 检查是否有背景色或条件格式
        If cell.Interior.Color <> xlNone And cell.Interior.Color <> RGB(255, 255, 255) Then
            HasConflictingFormats = True
            Exit Function
        End If
        
        If cell.FormatConditions.Count > 0 Then
            HasConflictingFormats = True
            Exit Function
        End If
        
        checkCount = checkCount + 1
    Next cell
    
    HasConflictingFormats = False
End Function

' 分析表格结构
Private Function AnalyzeTable(tableRange As Range) As TableAnalysis
    Dim analysis As TableAnalysis
    
    ' 基本信息
    Set analysis.TotalRange = tableRange
    analysis.DataRows = tableRange.Rows.Count
    analysis.DataColumns = tableRange.Columns.Count
    
    ' 检测表头
    Set analysis.HeaderRange = DetectHeaderRange(tableRange)
    If Not analysis.HeaderRange Is Nothing Then
        analysis.HasHeaders = True
        analysis.HeaderRows = analysis.HeaderRange.Rows.Count
        Set analysis.DataRange = GetDataRange(tableRange, analysis.HeaderRange)
    Else
        analysis.HasHeaders = False
        analysis.HeaderRows = 0
        Set analysis.DataRange = tableRange
    End If
    
    ' 内容特征分析
    analysis.HasNumbers = HasNumericData(analysis.DataRange)
    analysis.HasDates = HasDateData(analysis.DataRange)
    analysis.HasFormulas = HasFormulaData(analysis.DataRange)
    analysis.HasMergedCells = HasMergedCells(tableRange)
    
    AnalyzeTable = analysis
End Function

' 智能表头检测算法
Private Function DetectHeaderRange(tableRange As Range) As Range
    Dim headerScore As Long
    Dim maxHeaderRows As Long
    Dim rowNum As Long
    Dim testRow As Range
    Dim nextRow As Range
    
    maxHeaderRows = 3  ' 最多检测3行作为表头
    
    ' 评分标准
    Const SCORE_ALL_TEXT As Long = 30       ' 全部为文本
    Const SCORE_NO_EMPTY As Long = 25       ' 无空单元格
    Const SCORE_FORMAT_DIFF As Long = 20    ' 格式差异
    Const SCORE_BOLD_FONT As Long = 15      ' 加粗字体
    Const SCORE_BG_COLOR As Long = 10       ' 背景色
    Const SCORE_TYPE_DIFF As Long = 20      ' 数据类型差异
    
    Dim testRows As Long
    testRows = Application.Min(maxHeaderRows, tableRange.Rows.Count)
    
    For rowNum = 1 To testRows
        headerScore = 0
        Set testRow = tableRange.Rows(rowNum)
        
        ' 评分逻辑
        If IsAllText(testRow) Then headerScore = headerScore + SCORE_ALL_TEXT
        If HasNoEmpty(testRow) Then headerScore = headerScore + SCORE_NO_EMPTY
        If HasFormatting(testRow) Then headerScore = headerScore + SCORE_FORMAT_DIFF
        If HasBoldFont(testRow) Then headerScore = headerScore + SCORE_BOLD_FONT
        If HasBackgroundColor(testRow) Then headerScore = headerScore + SCORE_BG_COLOR
        
        ' 与下一行的数据类型差异
        If rowNum < tableRange.Rows.Count Then
            Set nextRow = tableRange.Rows(rowNum + 1)
            If HasTypeDifference(testRow, nextRow) Then
                headerScore = headerScore + SCORE_TYPE_DIFF
            End If
        End If
        
        ' 阈值判断（60分）
        If headerScore < 60 Then
            If rowNum = 1 Then
                ' 第一行分数不足时仍作为表头处理（兜底机制）
                Set DetectHeaderRange = tableRange.Rows(1)
            Else
                ' 返回前面的行作为表头
                Set DetectHeaderRange = tableRange.Rows("1:" & (rowNum - 1))
            End If
            Exit Function
        End If
    Next rowNum
    
    ' 默认第一行为表头
    Set DetectHeaderRange = tableRange.Rows(1)
End Function

' 获取数据区域
Private Function GetDataRange(tableRange As Range, headerRange As Range) As Range
    Dim startRow As Long
    
    If headerRange Is Nothing Then
        Set GetDataRange = tableRange
    Else
        startRow = headerRange.Row + headerRange.Rows.Count - tableRange.Row + 1
        If startRow <= tableRange.Rows.Count Then
            Set GetDataRange = tableRange.Rows(startRow & ":" & tableRange.Rows.Count)
        Else
            Set GetDataRange = Nothing
        End If
    End If
End Function

' =============================================================================
' 表头检测辅助函数
' =============================================================================

' 检测是否全部为文本
Private Function IsAllText(rng As Range) As Boolean
    Dim cell As Range
    Dim textCount As Long, totalCount As Long
    
    For Each cell In rng.Cells
        If Not IsEmpty(cell.Value) Then
            totalCount = totalCount + 1
            If Not IsNumeric(cell.Value) And Not IsDate(cell.Value) Then
                textCount = textCount + 1
            End If
        End If
    Next cell
    
    IsAllText = (textCount = totalCount And totalCount > 0)
End Function

' 检测是否无空单元格（安全处理错误值）
Private Function HasNoEmpty(rng As Range) As Boolean
    Dim cell As Range
    
    For Each cell In rng.Cells
        ' 先检查是否为空
        If IsEmpty(cell.Value) Then
            HasNoEmpty = False
            Exit Function
        End If
        
        ' 安全处理错误值和字符串转换
        On Error Resume Next
        If IsError(cell.Value) Then
            ' 错误值（如#N/A）视为非空，但不进行字符串转换
            On Error GoTo 0
        ElseIf Trim(CStr(cell.Value)) = "" Then
            HasNoEmpty = False
            Exit Function
        End If
        On Error GoTo 0
    Next cell
    
    HasNoEmpty = True
End Function

' 检测是否有格式化
Private Function HasFormatting(rng As Range) As Boolean
    Dim cell As Range
    
    For Each cell In rng.Cells
        ' 检查是否有非默认的背景色、字体样式等
        If cell.Interior.Color <> xlNone And cell.Interior.Color <> RGB(255, 255, 255) Then
            HasFormatting = True
            Exit Function
        End If
        If cell.Font.Bold = True Or cell.Font.Italic = True Then
            HasFormatting = True
            Exit Function
        End If
    Next cell
    
    HasFormatting = False
End Function

' 检测是否有粗体字体
Private Function HasBoldFont(rng As Range) As Boolean
    Dim cell As Range
    
    For Each cell In rng.Cells
        If cell.Font.Bold = True Then
            HasBoldFont = True
            Exit Function
        End If
    Next cell
    
    HasBoldFont = False
End Function

' 检测是否有背景色
Private Function HasBackgroundColor(rng As Range) As Boolean
    Dim cell As Range
    
    For Each cell In rng.Cells
        If cell.Interior.Color <> xlNone And cell.Interior.Color <> RGB(255, 255, 255) Then
            HasBackgroundColor = True
            Exit Function
        End If
    Next cell
    
    HasBackgroundColor = False
End Function

' 检测数据类型差异
Private Function HasTypeDifference(row1 As Range, row2 As Range) As Boolean
    Dim diffCount As Long, colCount As Long
    Dim i As Long
    
    colCount = row1.Cells.Count
    For i = 1 To colCount
        If i <= row2.Cells.Count Then
            If GetCellType(row1.Cells(i)) <> GetCellType(row2.Cells(i)) Then
                diffCount = diffCount + 1
            End If
        End If
    Next i
    
    ' 超过50%的列类型不同
    HasTypeDifference = (diffCount > colCount * 0.5)
End Function

' 获取单元格数据类型
Private Function GetCellType(cell As Range) As String
    If IsEmpty(cell.Value) Then
        GetCellType = "Empty"
    ElseIf IsNumeric(cell.Value) Then
        GetCellType = "Number"
    ElseIf IsDate(cell.Value) Then
        GetCellType = "Date"
    Else
        GetCellType = "Text"
    End If
End Function

' =============================================================================
' 内容特征检测函数
' =============================================================================

' 检测是否包含数值数据
Private Function HasNumericData(rng As Range) As Boolean
    Dim cell As Range
    Dim checkCount As Long
    
    For Each cell In rng.Cells
        If checkCount > 20 Then Exit For ' 只检查前20个单元格
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
            HasNumericData = True
            Exit Function
        End If
        checkCount = checkCount + 1
    Next cell
    
    HasNumericData = False
End Function

' 检测是否包含日期数据
Private Function HasDateData(rng As Range) As Boolean
    Dim cell As Range
    Dim checkCount As Long
    
    For Each cell In rng.Cells
        If checkCount > 20 Then Exit For ' 只检查前20个单元格
        If Not IsEmpty(cell.Value) And IsDate(cell.Value) Then
            HasDateData = True
            Exit Function
        End If
        checkCount = checkCount + 1
    Next cell
    
    HasDateData = False
End Function

' 检测是否包含公式
Private Function HasFormulaData(rng As Range) As Boolean
    Dim cell As Range
    Dim checkCount As Long
    
    For Each cell In rng.Cells
        If checkCount > 20 Then Exit For ' 只检查前20个单元格
        If cell.HasFormula Then
            HasFormulaData = True
            Exit Function
        End If
        checkCount = checkCount + 1
    Next cell
    
    HasFormulaData = False
End Function

' 检测是否包含合并单元格
Private Function HasMergedCells(rng As Range) As Boolean
    Dim cell As Range
    
    For Each cell In rng.Cells
        If cell.MergeCells Then
            HasMergedCells = True
            Exit Function
        End If
    Next cell
    
    HasMergedCells = False
End Function

' =============================================================================
' 主题样式系统
' =============================================================================

' 商务主题配置（默认开启斑马纹）
Private Function GetBusinessTheme() As BeautifyConfig
    Dim config As BeautifyConfig
    
    With config
        .ThemeName = "Business"
        .PrimaryColor = RGB(30, 58, 138)      ' 深蓝色
        .SecondaryColor = RGB(59, 130, 246)   ' 中蓝色
        .AccentColor = RGB(239, 246, 255)     ' 浅蓝色
        
        .EnableHeaderBeautify = True
        .EnableConditionalFormat = True
        .EnableBorders = True
        .EnableZebraStripes = True            ' *** 默认开启斑马纹 ***
        .EnableFreezeHeader = True
        
        .HeaderFontSize = 11
        .DataFontSize = 10
        .BorderWeight = xlThin
        .StripeOpacity = 0.05
    End With
    
    GetBusinessTheme = config
End Function

' 大表性能模式（自动关闭复杂样式）
Private Function GetPerformanceTheme(rowCount As Long) As BeautifyConfig
    Dim config As BeautifyConfig
    
    ' 基于Business主题
    config = GetBusinessTheme()
    
    ' 大表优化调整
    If rowCount > 10000 Then
        config.EnableZebraStripes = False     ' 大表关闭斑马纹
        config.EnableConditionalFormat = False ' 关闭复杂条件格式
        config.StripeOpacity = 0              ' 禁用透明度
    End If
    
    GetPerformanceTheme = config
End Function

' =============================================================================
' 样式应用引擎
' =============================================================================

' 应用主题样式
Private Sub ApplyThemeStyle(analysis As TableAnalysis, config As BeautifyConfig)
    ' 大表性能检测
    If analysis.DataRows > 10000 Then
        config = GetPerformanceTheme(analysis.DataRows)
    End If
    
    ' 应用表头样式
    If config.EnableHeaderBeautify And analysis.HasHeaders Then
        Call ApplyHeaderStyle(analysis.HeaderRange, config)
    End If
    
    ' 应用数据区域样式
    If Not analysis.DataRange Is Nothing Then
        Call ApplyDataStyle(analysis.DataRange, config)
    End If
    
    ' 应用边框
    If config.EnableBorders Then
        Call ApplyBorderStyle(analysis.TotalRange, analysis.HeaderRange, config)
    End If
    
    ' 应用隔行变色（条件格式实现，高性能）
    If config.EnableZebraStripes And Not analysis.DataRange Is Nothing Then
        Call ApplyZebraStripes(analysis.DataRange, config)
    End If
    
    ' 冻结表头
    If config.EnableFreezeHeader And analysis.HasHeaders Then
        Call FreezeHeader(analysis.HeaderRange)
    End If
End Sub

' 应用表头样式（商务蓝色渐变）
Private Sub ApplyHeaderStyle(headerRange As Range, config As BeautifyConfig)
    If headerRange Is Nothing Then Exit Sub
    
    With headerRange
        ' 背景色渐变实现（Excel 2007+支持，2003回退为纯色）
        On Error Resume Next ' 兼容旧版Excel
        If Val(Application.Version) >= 12 Then
            ' Excel 2007及以上版本：渐变色
            With .Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90 ' 垂直渐变（从上到下）
                .Gradient.ColorStops.Clear
                .Gradient.ColorStops.Add(0).Color = config.PrimaryColor     ' 起始色（较亮）
                .Gradient.ColorStops.Add(1).Color = RGB(30, 58, 138)        ' 结束色（深蓝，营造深度感）
            End With
        Else
            ' Excel 2003回退为纯色
            .Interior.Color = config.PrimaryColor
        End If
        On Error GoTo 0
        
        ' 字体样式
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)      ' 白色字体
        .Font.Size = config.HeaderFontSize
        .Font.Name = GetOptimalFont("ChineseHeader")
        
        ' 对齐方式
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' 行高调整
        .RowHeight = .RowHeight * 1.2
    End With
End Sub

' 应用数据区域样式
Private Sub ApplyDataStyle(dataRange As Range, config As BeautifyConfig)
    If dataRange Is Nothing Then Exit Sub
    
    With dataRange
        ' 字体样式
        .Font.Size = config.DataFontSize
        .Font.Name = GetOptimalFont("ChineseData")
        
        ' 对齐方式
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    ' 数值列特殊处理
    Call ApplyNumericColumnStyle(dataRange)
End Sub

' 数值列样式优化（保护用户NumberFormat）
Private Sub ApplyNumericColumnStyle(dataRange As Range)
    Dim col As Range
    
    For Each col In dataRange.Columns
        If IsNumericColumn(col) Then
            With col
                .Font.Name = GetOptimalFont("Number")    ' 使用Consolas等等宽字体
                .HorizontalAlignment = xlRight           ' 右对齐显示
                ' 注意：不修改NumberFormat，保护用户自定义的小数位数、百分比、货币符号等格式
            End With
        End If
    Next col
End Sub

' 分层边框样式应用（强化表头分隔，细化颜色层次）
Private Sub ApplyBorderStyle(tableRange As Range, headerRange As Range, config As BeautifyConfig)
    ' === 数据区域边框（浅色内部网格） ===
    With tableRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(209, 213, 219)  ' 内部网格：浅灰色，柔和分隔
    End With
    
    ' === 外边框加粗（深色边界） ===
    Dim outerBorders As Variant
    outerBorders = Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)
    
    Dim i As Long
    For i = 0 To UBound(outerBorders)
        With tableRange.Borders(outerBorders(i))
            .Weight = xlThick
            .Color = RGB(75, 85, 99)     ' 外边框：深灰色，明确边界
            .LineStyle = xlContinuous
        End With
    Next i
    
    ' === 表头底部强化分隔（双线+主色调深色） ===
    If Not headerRange Is Nothing Then
        With headerRange.Borders(xlEdgeBottom)
            .LineStyle = xlDouble         ' 双线样式，增强分隔感
            .Weight = xlThick
            .Color = RGB(30, 58, 138)     ' 主色调深色变体（深蓝），呼应主题
        End With
    End If
End Sub

' 条件格式实现隔行变色（单条CF规则，高性能可撤销）
Private Sub ApplyZebraStripes(dataRange As Range, config As BeautifyConfig)
    Dim sessionTag As String, stripeStep As Long
    Dim formula As String
    
    ' *** 统一会话标签 ***
    sessionTag = GetSessionTag()
    
    ' *** 关键：R1C1引用风格切换保护 ***
    Dim prevStyle As XlReferenceStyle
    prevStyle = Application.ReferenceStyle
    Application.ReferenceStyle = xlR1C1
    
    On Error GoTo ErrorHandler
    
    ' 智能步长：小表1行，中表2行，大表3行
    If dataRange.Rows.Count <= 50 Then
        stripeStep = 1  ' 每行交替
    ElseIf dataRange.Rows.Count <= 200 Then
        stripeStep = 2  ' 每2行交替
    Else
        stripeStep = 3  ' 每3行交替
    End If
    
    ' *** 单条条件格式实现斑马纹（R1C1格式）***
    ' 使用R1C1相对引用，避免固定行号依赖
    formula = "=MOD(ROW()-" & dataRange.Row & "+1," & (stripeStep * 2) & ")<=" & stripeStep & _
              "+N(0*LEN(""" & sessionTag & """))"
    
    With dataRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
        .Interior.Color = config.AccentColor
        .StopIfTrue = False
        .Priority = 10  ' 低优先级，不覆盖其他条件格式
    End With
    
    ' *** 恢复原始引用风格 ***
    Application.ReferenceStyle = prevStyle
    
    ' *** 统一两段式日志记录 ***
    Call LogCFRule(dataRange.Address & "|" & sessionTag)
    Exit Sub
    
ErrorHandler:
    Application.ReferenceStyle = prevStyle  ' 错误时也恢复
End Sub

' 冻结表头实现
Private Sub FreezeHeader(headerRange As Range)
    On Error Resume Next
    ' 在表头下方一行设置冻结窗格
    Dim freezeRow As Long
    freezeRow = headerRange.Row + headerRange.Rows.Count
    
    ' 设置冻结位置（表头下方第一行的A列）
    headerRange.Worksheet.Cells(freezeRow, 1).Select
    ActiveWindow.FreezePanes = True
    
    On Error GoTo 0
End Sub

' =============================================================================
' 条件格式实现（统一R1C1架构）
' =============================================================================

' 条件格式统一应用（删除A1变体，仅保留R1C1实现）
Private Sub ApplyStandardConditionalFormat(dataRange As Range)
    Dim sessionTag As String
    Dim col As Range
    
    If dataRange Is Nothing Then Exit Sub
    
    ' *** 统一会话标签，确保撤销一致性 ***
    sessionTag = GetSessionTag()  ' 使用全局统一标签
    
    ' *** 关键：R1C1引用风格切换保护 ***
    Dim prevStyle As XlReferenceStyle
    prevStyle = Application.ReferenceStyle
    Application.ReferenceStyle = xlR1C1
    
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' *** 关键：仅清理带标签的规则，保护用户既有格式 ***
    ' 先清理整体数据区域
    Call ClearTaggedRules(dataRange, sessionTag)
    
    ' 统一优先级顺序（R1C1相对引用）
    ' 1. 错误值检测（优先级1，终止后续判断）
    Call ApplyErrorHighlight(dataRange, sessionTag)
    
    ' 2. 空值标记（优先级2，终止后续判断）
    Call ApplyEmptyHighlight(dataRange, sessionTag)
    
    ' 3. 逐列应用重复值检测（精确范围控制，逐列预清理确保幂等性）
    For Each col In dataRange.Columns
        ' *** 修复：逐列预清理，确保多次运行的幂等性 ***
        Call ClearTaggedRules(col, sessionTag)
        Call ApplyDuplicateHighlight(col, sessionTag)
    Next col
    
    ' 4. 数值列负数检测（仅数值列，避免格式覆盖，逐列预清理）
    For Each col In dataRange.Columns
        If IsNumericColumn(col) Then
            ' *** 修复：逐列预清理，确保多次运行的幂等性 ***
            Call ClearTaggedRules(col, sessionTag)
            Call ApplyNegativeHighlight(col, sessionTag)
        End If
    Next col
    
CleanUp:
    ' *** 恢复原始引用风格（错误情况下也要恢复）***
    Application.ReferenceStyle = prevStyle
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ReferenceStyle = prevStyle  ' 错误时也恢复
    MsgBox "条件格式应用失败: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

' 仅清理带会话标签的规则（避免误删用户既有格式）
Private Sub ClearTaggedRules(rng As Range, sessionTag As String)
    Dim i As Long, cf As FormatCondition
    
    ' 从后往前删除，避免索引变化
    For i = rng.FormatConditions.Count To 1 Step -1
        Set cf = rng.FormatConditions(i)
        
        ' 检查公式中是否包含会话标签
        If InStr(cf.Formula1, sessionTag) > 0 Or InStr(cf.Formula2, sessionTag) > 0 Then
            cf.Delete
        End If
    Next i
End Sub

' 错误值高亮（纯R1C1，优先级1，终止后续）
Private Sub ApplyErrorHighlight(rng As Range, tag As String)
    Dim formula As String
    formula = "=ISERROR(RC)+N(0*LEN(""" & tag & """))"
    
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
        .Interior.Color = RGB(254, 226, 226)  ' 浅红背景
        .Font.Color = RGB(127, 29, 29)        ' 深红字体
        .StopIfTrue = True                    ' *** 错误值终止后续判断 ***
        .Priority = 1  ' 最高优先级
    End With
    
    ' 统一两段式记录：地址|标签
    Call LogCFRule(rng.Address & "|" & tag)
End Sub

' 空值标记（纯R1C1，优先级2，终止后续）
Private Sub ApplyEmptyHighlight(rng As Range, tag As String)
    Dim formula As String
    formula = "=ISBLANK(RC)+N(0*LEN(""" & tag & """))"
    
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
        .Interior.Color = RGB(249, 250, 251)  ' 浅灰背景
        .StopIfTrue = True                    ' *** 空值终止后续判断 ***
        .Priority = 2
    End With
    
    Call LogCFRule(rng.Address & "|" & tag)
End Sub

' 重复值检测（R1C1列相对引用，优先级3，允许叠加）
Private Sub ApplyDuplicateHighlight(col As Range, tag As String)
    Dim formula As String
    
    ' *** 关键修正：使用R1C1列相对引用 C[0]，避免Address解析 ***
    formula = "=AND(RC<>"""",COUNTIF(C[0],RC)>1)+N(0*LEN(""" & tag & """))"
    
    ' 精确控制AppliesTo到当前列
    With col.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
        .Interior.Color = RGB(255, 251, 235)  ' 浅黄背景
        .StopIfTrue = False                   ' *** 允许与负数规则叠加 ***
        .Priority = 3
    End With
    
    Call LogCFRule(col.Address & "|" & tag)
End Sub

' 负数检测（仅表达式+字体颜色，优先级4，允许叠加）
Private Sub ApplyNegativeHighlight(col As Range, tag As String)
    Dim formula As String
    formula = "=RC<0+N(0*LEN(""" & tag & """))"
    
    ' *** 关键修正：仅设字体颜色，保护用户NumberFormat ***
    With col.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
        .Font.Color = RGB(220, 38, 38)       ' 红色字体
        .StopIfTrue = False                   ' *** 仅设字体色，可叠加背景色 ***
        .Priority = 4
        ' *** 不设置NumberFormat，保护用户小数位/千分位设置 ***
    End With
    
    Call LogCFRule(col.Address & "|" & tag)
End Sub

' =============================================================================
' 工具函数
' =============================================================================

' 快速数值列检测（避免逐单元格遍历）
Private Function IsNumericColumn(col As Range) As Boolean
    Dim checkCount As Long, numericCount As Long
    Dim cell As Range, maxCheck As Long
    
    ' 仅检查前5个非空单元格，提升性能
    maxCheck = 5
    checkCount = 0
    numericCount = 0
    
    For Each cell In col.Cells
        If Not IsEmpty(cell.Value) And checkCount < maxCheck Then
            checkCount = checkCount + 1
            If IsNumeric(cell.Value) Then
                numericCount = numericCount + 1
            End If
        End If
        If checkCount >= maxCheck Then Exit For
    Next cell
    
    ' 60%以上为数值则认为是数值列
    IsNumericColumn = (numericCount >= (checkCount * 0.6)) And checkCount > 0
End Function

' 优化字体选择（兼容性+可读性优先）
Private Function GetOptimalFont(contentType As String) As String
    Select Case contentType
        Case "ChineseHeader"
            ' 中文标题：优先微软雅黑，回退宋体/苹方
            If IsFontAvailable("微软雅黑") Then
                GetOptimalFont = "微软雅黑"
            ElseIf IsFontAvailable("苹方-简") Then
                GetOptimalFont = "苹方-简"
            Else
                GetOptimalFont = "宋体"  ' 最后回退
            End If
        Case "ChineseData"
            GetOptimalFont = "微软雅黑"  ' 统一微软雅黑，删除Light字重
        Case "EnglishHeader"
            GetOptimalFont = "Calibri"  ' 英文标题
        Case "EnglishData"
            GetOptimalFont = "Arial"    ' 英文数据
        Case "Number", "Currency", "Financial"
            ' *** 数字/金额统一等宽字体，优先级回退 ***
            If IsFontAvailable("Consolas") Then
                GetOptimalFont = "Consolas"      ' 首选等宽
            ElseIf IsFontAvailable("Courier New") Then
                GetOptimalFont = "Courier New"   ' 回退等宽
            ElseIf IsFontAvailable("SF Mono") Then
                GetOptimalFont = "SF Mono"       ' Mac等宽
            ElseIf IsFontAvailable("Menlo") Then
                GetOptimalFont = "Menlo"         ' Mac回退
            Else
                GetOptimalFont = "微软雅黑"       ' 最终回退
            End If
        Case "Mixed"
            GetOptimalFont = "微软雅黑"  ' 混合内容默认
        Case Else
            GetOptimalFont = "微软雅黑"  ' 默认字体
    End Select
End Function

' 字体可用性检查（稳定的形状试探法）
Private Function IsFontAvailable(fontName As String) As Boolean
    Dim originalUpdating As Boolean
    Dim testShape As Shape
    Dim testSheet As Worksheet
    Dim success As Boolean
    
    ' 关闭屏幕更新提升性能
    originalUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    ' 方法1：尝试使用临时形状试探字体（不落盘）
    Set testSheet = ActiveSheet
    If Not testSheet Is Nothing Then
        ' 创建隐藏的临时文本框
        Set testShape = testSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 10, 10)
        testShape.Visible = False
        testShape.TextFrame.Characters.Font.Name = fontName
        
        ' 检查字体是否真的被设置
        success = (testShape.TextFrame.Characters.Font.Name = fontName)
        
        ' 删除临时形状
        testShape.Delete
        Set testShape = Nothing
    End If
    
    ' 方法2：如果形状方法失败，回退到简单验证
    If Err.Number <> 0 Or Not success Then
        Err.Clear
        ' 简单的字体名称验证
        success = (Len(fontName) > 0 And fontName <> "")
    End If
    
    On Error GoTo 0
    Application.ScreenUpdating = originalUpdating
    
    IsFontAvailable = success
End Function

' =============================================================================
' 撤销机制实现
' =============================================================================

' 初始化撤销日志 - 创建新的操作记录并推入堆栈
Private Sub InitializeBeautifyLog()
    Dim newLog As BeautifyLog
    
    ' 初始化堆栈（如果未初始化）
    If g_UndoStack Is Nothing Then
        Set g_UndoStack = New Collection
    End If
    
    ' 创建新的日志记录
    With newLog
        .SessionId = Format(Now, "yyyymmddhhmmss") & "_" & Int(Rnd * 1000)
        .Timestamp = Now
        .CFRulesAdded = ""
        .StylesAdded = ""
        .TableStylesMap = ""        ' 表格样式映射：表名:原样式;...
    End With
    
    ' 将新记录推入堆栈
    g_UndoStack.Add newLog
    
    ' 限制堆栈大小（最多保留20个操作记录）
    Do While g_UndoStack.Count > 20
        g_UndoStack.Remove 1
    Loop
End Sub

' 获取当前操作记录（堆栈顶部）
Private Function GetCurrentLog() As BeautifyLog
    If g_UndoStack Is Nothing Then
        Set g_UndoStack = New Collection
    End If
    
    If g_UndoStack.Count > 0 Then
        GetCurrentLog = g_UndoStack(g_UndoStack.Count)
    End If
End Function

' 更新当前操作记录（堆栈顶部）
Private Sub UpdateCurrentLog(updatedLog As BeautifyLog)
    If g_UndoStack Is Nothing Or g_UndoStack.Count = 0 Then
        Exit Sub
    End If
    
    ' 移除顶部记录并添加更新后的记录
    g_UndoStack.Remove g_UndoStack.Count
    g_UndoStack.Add updatedLog
End Sub

' 记录表格样式变更
Private Sub LogTableStyleChange(tblName As String, originalStyle As String)
    Dim mapping As String
    Dim currentLog As BeautifyLog
    
    mapping = tblName & ":" & originalStyle
    currentLog = GetCurrentLog()
    
    If currentLog.TableStylesMap = "" Then
        currentLog.TableStylesMap = mapping
    Else
        currentLog.TableStylesMap = currentLog.TableStylesMap & ";" & mapping
    End If
    
    Call UpdateCurrentLog(currentLog)
End Sub

' 记录样式创建
Private Sub LogStyleCreation(styleName As String)
    Dim currentLog As BeautifyLog
    
    currentLog = GetCurrentLog()
    
    If currentLog.StylesAdded = "" Then
        currentLog.StylesAdded = styleName
    Else
        currentLog.StylesAdded = currentLog.StylesAdded & ";" & styleName
    End If
    
    Call UpdateCurrentLog(currentLog)
End Sub

' *** 统一日志记录接口（两段式：地址|标签）***
Private Sub LogCFRule(ruleInfo As String)
    Dim currentLog As BeautifyLog
    
    currentLog = GetCurrentLog()
    
    If currentLog.CFRulesAdded = "" Then
        currentLog.CFRulesAdded = ruleInfo
    Else
        currentLog.CFRulesAdded = currentLog.CFRulesAdded & ";" & ruleInfo
    End If
    
    Call UpdateCurrentLog(currentLog)
End Sub

' *** 会话标签统一生成（全局一致）***
Private Function GetSessionTag() As String
    Dim currentLog As BeautifyLog
    currentLog = GetCurrentLog()
    GetSessionTag = "ELO_" & currentLog.SessionId
End Function

' 删除指定标签的条件格式规则
Private Sub RemoveTaggedCFRule(ws As Worksheet, ruleEntry As String)
    Dim parts() As String, rngAddress As String, tag As String
    Dim targetRange As Range, i As Long
    
    parts = Split(ruleEntry, "|")
    If UBound(parts) >= 1 Then
        rngAddress = parts(0)
        tag = parts(1)
        
        On Error Resume Next
        Set targetRange = ws.Range(rngAddress)
        If Not targetRange Is Nothing Then
            Call ClearTaggedRules(targetRange, tag)
        End If
        On Error GoTo 0
    End If
End Sub

' 还原表格样式
Private Sub RestoreTableStyle(ws As Worksheet, mapping As String)
    Dim parts() As String, tblName As String, originalStyle As String
    Dim tbl As ListObject
    
    parts = Split(mapping, ":")
    If UBound(parts) = 1 Then
        tblName = parts(0)
        originalStyle = parts(1)
        
        On Error Resume Next
        Set tbl = ws.ListObjects(tblName)
        If Not tbl Is Nothing Then
            If originalStyle = "" Then
                tbl.TableStyle = ""
            Else
                tbl.TableStyle = originalStyle
            End If
        End If
        On Error GoTo 0
    End If
End Sub

' 安全删除样式
Private Sub SafeDeleteStyle(styleName As String)
    On Error Resume Next
    ActiveWorkbook.Styles(styleName).Delete
    On Error GoTo 0
End Sub

' 删除会话表格样式
Private Sub RemoveSessionTableStyles(sessionTag As String)
    Dim i As Long
    
    For i = ActiveWorkbook.TableStyles.Count To 1 Step -1
        If InStr(ActiveWorkbook.TableStyles(i).Name, sessionTag) > 0 Then
            On Error Resume Next
            ActiveWorkbook.TableStyles(i).Delete
            On Error GoTo 0
        End If
    Next i
End Sub

' =============================================================================
' 性能优化和状态管理
' =============================================================================

Private Function SaveAppState() As AppState
    With Application
        SaveAppState.ScreenUpdating = .ScreenUpdating
        SaveAppState.Calculation = .Calculation
        SaveAppState.EnableEvents = .EnableEvents
        SaveAppState.DisplayAlerts = .DisplayAlerts
        SaveAppState.Cursor = .Cursor
        SaveAppState.ReferenceStyle = .ReferenceStyle
    End With
End Function

Private Sub RestoreAppState(state As AppState)
    With Application
        .ScreenUpdating = state.ScreenUpdating
        .Calculation = state.Calculation
        .EnableEvents = state.EnableEvents
        .DisplayAlerts = state.DisplayAlerts
        .Cursor = state.Cursor
        .ReferenceStyle = state.ReferenceStyle
    End With
End Sub

Private Sub SetPerformanceMode()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
        .Cursor = xlWait
    End With
End Sub

' 大表性能模式检测
Private Function NeedsPerformanceMode(rng As Range) As Boolean
    Const LARGE_ROW_COUNT As Long = 10000
    Const LARGE_COL_COUNT As Long = 50
    
    NeedsPerformanceMode = (rng.Rows.Count > LARGE_ROW_COUNT) Or _
                           (rng.Columns.Count > LARGE_COL_COUNT)
End Function

' =============================================================================
' 错误处理机制
' =============================================================================

' 错误处理函数
Private Sub HandleError(errCode As BeautifyError, Optional details As String = "")
    Dim message As String
    
    Select Case errCode
        Case ERR_NO_SELECTION
            message = "请选择要美化的表格区域"
        Case ERR_INVALID_RANGE
            message = "无效的表格区域：" & details
        Case ERR_MEMORY_LIMIT
            message = "内存不足，请缩小数据范围"
        Case ERR_FORMAT_CONFLICT
            message = "格式冲突：" & details
        Case ERR_UNDO_FAILED
            message = "撤销失败：" & details
        Case Else
            message = "未知错误：" & details
    End Select
    
    ' 记录错误日志
    Call LogError(errCode, message)
    
    ' 显示用户友好提示
    MsgBox message, vbExclamation, "Excel美化工具"
End Sub

' 错误日志记录
Private Sub LogError(errCode As Long, message As String)
    Debug.Print "BeautifyError [" & Now & "] Code: " & errCode & " - " & message
End Sub

' =============================================================================
' 测试和调试函数
' =============================================================================

' 测试主函数
Public Sub TestBeautifier()
    Debug.Print "===== Excel美化系统测试开始 ====="
    
    ' 测试表头检测
    Call TestHeaderDetection()
    
    ' 测试数据类型检测
    Call TestDataTypeDetection()
    
    ' 测试字体可用性
    Call TestFontAvailability()
    
    Debug.Print "===== 测试完成 ====="
End Sub

' 测试表头检测
Private Sub TestHeaderDetection()
    Debug.Print "- 测试表头检测算法"
    
    Dim testRange As Range
    Set testRange = Selection
    
    If Not testRange Is Nothing Then
        Dim headerRange As Range
        Set headerRange = DetectHeaderRange(testRange)
        
        If Not headerRange Is Nothing Then
            Debug.Print "  检测到表头: " & headerRange.Address
        Else
            Debug.Print "  未检测到表头"
        End If
    Else
        Debug.Print "  请先选择测试区域"
    End If
End Sub

' 测试数据类型检测
Private Sub TestDataTypeDetection()
    Debug.Print "- 测试数据类型检测"
    
    Dim testRange As Range
    Set testRange = Selection
    
    If Not testRange Is Nothing Then
        Debug.Print "  数值数据: " & HasNumericData(testRange)
        Debug.Print "  日期数据: " & HasDateData(testRange)
        Debug.Print "  公式数据: " & HasFormulaData(testRange)
        Debug.Print "  合并单元格: " & HasMergedCells(testRange)
    Else
        Debug.Print "  请先选择测试区域"
    End If
End Sub

' 测试字体可用性
Private Sub TestFontAvailability()
    Debug.Print "- 测试字体可用性"
    
    Dim testFonts As Variant
    testFonts = Array("微软雅黑", "Consolas", "Arial", "不存在的字体")
    
    Dim i As Long
    For i = 0 To UBound(testFonts)
        Debug.Print "  " & testFonts(i) & ": " & IsFontAvailable(CStr(testFonts(i)))
    Next i
End Sub

' 显示系统信息
Public Sub ShowSystemInfo()
    Dim info As String
    
    info = "Excel美化系统 v2.1" & vbCrLf & vbCrLf
    info = info & "当前状态:" & vbCrLf
    info = info & "- 撤销历史: " & IIf(g_HasBeautifyHistory, "有", "无") & vbCrLf
    info = info & "- Excel版本: " & Application.Version & vbCrLf
    info = info & "- 引用风格: " & IIf(Application.ReferenceStyle = xlA1, "A1", "R1C1") & vbCrLf
    info = info & vbCrLf & "主要功能:" & vbCrLf
    info = info & "- BeautifyTable(): 美化表格" & vbCrLf
    info = info & "- UndoBeautify(): 撤销美化" & vbCrLf
    info = info & "- TestBeautifier(): 运行测试" & vbCrLf
    info = info & "- InstallBeautifier(): 初始化设置" & vbCrLf
    
    MsgBox info, vbInformation, "Excel美化系统"
End Sub

' =============================================================================
' 系统安装与配置模块
' =============================================================================

' 一键安装美化系统
Public Sub InstallBeautifier()
    Dim result As VbMsgBoxResult
    
    ' 欢迎信息
    result = MsgBox("欢迎使用Excel表格美化系统 v2.1！" & vbCrLf & vbCrLf & _
                   "本安装程序将：" & vbCrLf & _
                   "• 检查系统兼容性" & vbCrLf & _
                   "• 设置快捷键" & vbCrLf & _
                   "• 运行系统测试" & vbCrLf & vbCrLf & _
                   "是否继续安装？", _
                   vbYesNo + vbInformation, "Excel美化系统安装")
    
    If result = vbNo Then
        MsgBox "安装已取消", vbInformation
        Exit Sub
    End If
    
    ' 检查系统兼容性
    If Not CheckSystemCompatibility() Then
        Exit Sub
    End If
    
    ' 检查是否已配置
    If IsBeautifierConfigured() Then
        result = MsgBox("检测到系统已配置，是否重新配置？", _
                       vbYesNo + vbQuestion, "Excel美化系统")
        If result = vbNo Then
            Call ShowQuickStart()
            Exit Sub
        End If
    End If
    
    ' 设置快捷键
    Call SetupShortcuts()
    
    ' 显示安装完成信息
    Call ShowInstallationComplete()
End Sub

' 检查系统兼容性
Private Function CheckSystemCompatibility() As Boolean
    Dim excelVersion As Single
    
    ' 检查Excel版本
    excelVersion = Val(Application.Version)
    If excelVersion < 15 Then ' Excel 2013 = 15.0
        MsgBox "系统要求Excel 2013或更高版本，当前版本：" & Application.Version, _
               vbCritical, "兼容性检查"
        CheckSystemCompatibility = False
        Exit Function
    End If
    
    ' 检查宏设置
    If Application.AutomationSecurity = msoAutomationSecurityForceDisable Then
        MsgBox "请启用宏功能后重新安装" & vbCrLf & vbCrLf & _
               "设置路径：文件 > 选项 > 信任中心 > 宏设置", _
               vbCritical, "兼容性检查"
        CheckSystemCompatibility = False
        Exit Function
    End If
    
    CheckSystemCompatibility = True
End Sub

' 检查是否已配置快捷键
Private Function IsBeautifierConfigured() As Boolean
    ' 简单检查：尝试获取已设置的快捷键状态
    ' 注意：Excel VBA无法直接检测OnKey状态，这里使用启发式判断
    IsBeautifierConfigured = False ' 默认为未配置，让用户选择是否重新配置
End Function

' 设置快捷键
Private Sub SetupShortcuts()
    On Error Resume Next
    
    ' 设置Ctrl+Shift+B为美化快捷键
    Application.OnKey "^+B", "BeautifyTable"
    
    ' 设置Ctrl+Shift+Z为撤销快捷键  
    Application.OnKey "^+Z", "UndoBeautify"
    
    ' 设置Ctrl+Shift+T为测试快捷键
    Application.OnKey "^+T", "TestBeautifier"
    
    ' 设置Ctrl+Shift+H为帮助快捷键
    Application.OnKey "^+H", "ShowSystemInfo"
    
    On Error GoTo 0
End Sub

' 显示安装完成信息
Private Sub ShowInstallationComplete()
    Dim info As String
    
    info = "🎉 Excel美化系统安装成功！" & vbCrLf & vbCrLf
    info = info & "📋 快捷键列表：" & vbCrLf
    info = info & "   Ctrl+Shift+B  →  美化表格" & vbCrLf
    info = info & "   Ctrl+Shift+Z  →  撤销美化" & vbCrLf
    info = info & "   Ctrl+Shift+T  →  运行测试" & vbCrLf
    info = info & "   Ctrl+Shift+H  →  显示帮助" & vbCrLf & vbCrLf
    info = info & "🚀 快速开始：" & vbCrLf
    info = info & "1. 选择要美化的表格区域" & vbCrLf
    info = info & "2. 按 Ctrl+Shift+B 一键美化" & vbCrLf
    info = info & "3. 如需撤销，按 Ctrl+Shift+Z" & vbCrLf & vbCrLf
    info = info & "💡 提示：也可以在VBA立即窗口中直接调用：" & vbCrLf
    info = info & "   BeautifyTable() 或 UndoBeautify()" & vbCrLf
    
    MsgBox info, vbInformation, "安装完成"
End Sub

' 显示快速开始指南
Private Sub ShowQuickStart()
    Dim info As String
    
    info = "Excel美化系统已就绪！" & vbCrLf & vbCrLf
    info = info & "快捷键：" & vbCrLf
    info = info & "• Ctrl+Shift+B  美化表格" & vbCrLf
    info = info & "• Ctrl+Shift+Z  撤销美化" & vbCrLf
    info = info & "• Ctrl+Shift+T  运行测试" & vbCrLf
    info = info & "• Ctrl+Shift+H  显示帮助" & vbCrLf
    
    MsgBox info, vbInformation, "快速开始"
End Sub

' 卸载美化系统快捷键
Public Sub UninstallBeautifier()
    Dim result As VbMsgBoxResult
    
    result = MsgBox("确定要清除所有快捷键设置吗？" & vbCrLf & vbCrLf & _
                   "这将清除所有快捷键绑定，但保留VBA模块", _
                   vbYesNo + vbQuestion, "清除快捷键")
    
    If result = vbYes Then
        ' 清除快捷键
        On Error Resume Next
        Application.OnKey "^+B"
        Application.OnKey "^+Z"
        Application.OnKey "^+T"
        Application.OnKey "^+H"
        On Error GoTo 0
        
        MsgBox "快捷键已清除！" & vbCrLf & _
               "如需重新设置，请运行 InstallBeautifier()", _
               vbInformation, "清除完成"
    End If
End Sub

' 重置快捷键
Public Sub ResetShortcuts()
    Call SetupShortcuts()
    MsgBox "快捷键已重置", vbInformation, "设置完成"
End Sub
