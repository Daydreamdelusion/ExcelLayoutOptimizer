# Excel布局优化系统 - 使用示例

## 🎯 标题优先功能演示

### 示例1：长标题自动换行

**原始情况：**
```
| 客户名称 | 产品销售金额（... | 去年同期对比... |
| 张三     | 12345.67        | 15%             |
```

**优化后：**
```
| 客户名称 | 产品销售金额      | 去年同期对比     |
|         | （万元）          | 增长率（%）      |
| 张三     | 12345.67        | 15%             |
```

### 示例2：标题与数据宽度平衡

**原始情况：**
```
| ID | Name | Very Long Description Column |
| 1  | A    | Short text                   |
| 2  | B    | Another                      |
```

**优化后：**
```
| ID | Name | Very Long Description |
|    |      | Column                |
| 1  | A    | Short text           |
| 2  | B    | Another              |
```

## 📝 操作步骤详解

### 基本操作流程

1. **选择数据区域**
   ```
   按住鼠标左键拖拽，选中包含表头的完整数据区域
   ```

2. **运行优化命令**
   ```vba
   ' 方法1：直接调用
   Call OptimizeLayout
   
   ' 方法2：快捷键（需先安装）
   Ctrl + Shift + L
   ```

3. **确认预览信息**
   ```
   [预览对话框]
   优化区域: $A$1:$D$100
   ------------------------------
   • 总列数: 4
   • 需调整: 3 列
   • 需换行: 1 列（包含标题）
   • 宽度范围: 8.4 - 45.0
   • 标题优先: 已启用
   ------------------------------
   预计耗时: 1.2 秒
   
   是否继续？
   ```

4. **查看优化结果**
   ```
   [完成提示]
   优化完成！
   - 已优化：4列
   - 用时：0.8秒
   提示：按Ctrl+Z可撤销本次操作
   ```

### 高级配置示例

#### 自定义标题优先设置
```vba
Sub CustomHeaderPriorityDemo()
    ' 设置标题优先模式
    Call OptimizeLayout  ' 会弹出配置对话框
    
    ' 在配置对话框中：
    ' 1. 最大列宽设为 40
    ' 2. 启用标题优先模式
    ' 3. 确认优化
End Sub
```

#### 批量处理多个工作表
```vba
Sub BatchOptimizeSheets()
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        If ws.UsedRange.Rows.Count > 1 Then
            ws.UsedRange.Select
            Call QuickOptimize  ' 快速优化，跳过预览
        End If
    Next ws
End Sub
```

## 🔧 常见问题处理

### Q1：标题过长导致行高过高
**问题**：标题换行后第一行变得很高
**解决**：调整标题最大换行行数
```vba
' 修改配置限制标题最多换2行
g_Config.HeaderMaxWrapLines = 2
```

### Q2：标题优先模式不生效
**问题**：设置了标题优先但效果不明显
**检查**：
1. 确认启用了智能表头识别
2. 检查第一行是否被正确识别为标题
3. 验证配置参数是否正确

```vba
' 检查配置
Debug.Print g_Config.HeaderPriority        ' 应该是 True
Debug.Print g_Config.SmartHeaderDetection  ' 应该是 True
```

### Q3：某些列没有被优化
**原因**：可能包含合并单元格
**解决**：
1. 先取消合并单元格
2. 然后运行优化
3. 如需要再重新合并

## 💡 最佳实践建议

### 数据准备
1. **清理数据**：删除空行和无关内容
2. **标准化格式**：统一日期、数字格式
3. **检查合并**：处理合并单元格

### 优化策略
1. **先预览**：大型表格先用预览模式检查
2. **分段处理**：超大表格可分区域处理
3. **保存备份**：重要数据先备份再优化

### 参数调整
```vba
' 针对不同场景的推荐配置

' 报表场景（标题重要）
g_Config.HeaderPriority = True
g_Config.MaxColumnWidth = 40
g_Config.HeaderMaxWrapLines = 2

' 数据输入场景（数据重要）
g_Config.HeaderPriority = False
g_Config.MaxColumnWidth = 60
g_Config.TextBuffer = 3

' 演示场景（美观重要）
g_Config.HeaderPriority = True
g_Config.MaxColumnWidth = 35
g_Config.HeaderMinHeight = 20
```

## 🧪 测试数据生成

### 创建测试表格
```vba
Sub CreateTestData()
    ' 清空当前工作表
    Cells.Clear
    
    ' 创建测试标题（包含长标题）
    Range("A1").Value = "客户编号"
    Range("B1").Value = "客户全称（包含分公司详细信息）"
    Range("C1").Value = "产品销售金额（万元）"
    Range("D1").Value = "去年同期对比增长率（百分比）"
    
    ' 创建测试数据
    Dim i As Long
    For i = 2 To 21
        Range("A" & i).Value = "C00" & (i - 1)
        Range("B" & i).Value = "客户" & (i - 1) & "有限公司"
        Range("C" & i).Value = Round(Rnd() * 10000, 2)
        Range("D" & i).Value = Round((Rnd() - 0.5) * 100, 1) & "%"
    Next i
    
    MsgBox "测试数据已创建，可以开始测试优化功能！"
End Sub
```

### 运行完整测试
```vba
Sub FullFunctionTest()
    ' 创建测试数据
    Call CreateTestData
    
    ' 选择数据区域
    Range("A1:D21").Select
    
    ' 显示优化前状态
    MsgBox "优化前的表格状态，请查看列宽和标题显示"
    
    ' 运行优化
    Call OptimizeLayout
    
    ' 显示优化后状态
    MsgBox "优化后的表格状态，注意标题和列宽的变化"
    
    ' 测试撤销功能
    If MsgBox("是否测试撤销功能？", vbYesNo) = vbYes Then
        Call UndoLastOptimization
        MsgBox "已撤销优化，表格恢复原状"
    End If
End Sub
```

---

**提示**：建议在使用前先运行 `Call RunTestSuite` 验证系统功能正常。
