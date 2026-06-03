# COM 接口与 VBA 标准文档

## 什么是 COM？

**COM (Component Object Model)** 是微软的二进制接口标准，允许不同语言（C#、Python、VBA、C++、JavaScript）通过统一接口调用同一个对象。

```
┌──────────────┐    COM 接口    ┌──────────────────────┐
│  VBA 宏      │───────────────▶│                      │
│  Python      │───────────────▶│  PowerPoint.exe      │
│  C# .NET     │───────────────▶│  (COM Server)        │
│  C++         │───────────────▶│                      │
│  PowerShell  │───────────────▶│                      │
└──────────────┘                └──────────────────────┘
```

核心概念：
- **COM Server**: 提供功能的程序（如 PowerPoint.exe）
- **COM Client**: 调用功能的程序（我们的编辑器）
- **IDispatch**: 晚绑定接口，VBA/Python 通过它调用方法
- **ProgID**: 程序标识符，如 `"PowerPoint.Application"`
- **CLSID**: 全局唯一类标识符，注册在 Windows 注册表

## PowerPoint COM 对象模型

微软 Office 的 COM 对象模型层次结构：

```
Application                          ← PowerPoint.Application
├── Presentations                    ← 所有打开的演示文稿
│   └── Presentation                 ← 单个 PPTX 文件
│       ├── PageSetup                ← 页面设置 (宽/高)
│       ├── Slides                   ← 所有幻灯片
│       │   └── Slide                ← 单张幻灯片
│       │       ├── Shapes           ← 所有元素
│       │       │   └── Shape        ← 单个元素
│       │       │       ├── TextFrame
│       │       │       │   └── TextRange
│       │       │       │       ├── Text
│       │       │       │       ├── Font
│       │       │       │       ├── ParagraphFormat
│       │       │       │       ├── Paragraphs()
│       │       │       │       └── Characters()
│       │       │       ├── Fill
│       │       │       ├── Line
│       │       │       ├── Table
│       │       │       │   ├── Rows / Columns
│       │       │       │   └── Cell(row, col)
│       │       │       ├── PlaceholderFormat
│       │       │       └── HasTextFrame / HasTable
│       │       ├── SlideShowTransition
│       │       ├── TimeLine
│       │       │   └── MainSequence
│       │       │       └── Effect
│       │       └── Export()
│       └── SaveAs()
└── Quit()
```

## VBA vs COM：同一套接口

VBA 宏和外部 COM 调用**完全等价**，用的是同一套接口。

### VBA 写法

```vba
Sub ModifyTitle()
    Dim ppt As Presentation
    Set ppt = Application.Presentations.Open("C:\test.pptx")
    
    Dim slide As slide
    Set slide = ppt.Slides(1)
    
    Dim shape As shape
    Set shape = slide.Shapes(1)
    
    shape.TextFrame.TextRange.Text = "新标题"
    shape.TextFrame.TextRange.Font.Size = 40
    shape.TextFrame.TextRange.Font.Bold = msoTrue
    shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 0, 0)
    
    ppt.SaveAs "C:\test_modified.pptx"
    ppt.Close
End Sub
```

### Python COM 写法 (pywin32)

```python
import win32com.client
import pythoncom

pythoncom.CoInitialize()
app = win32com.client.Dispatch("PowerPoint.Application")
ppt = app.Presentations.Open(r"C:\test.pptx", False, False, False)

slide = ppt.Slides(1)
shape = slide.Shapes(1)

shape.TextFrame.TextRange.Text = "新标题"
shape.TextFrame.TextRange.Font.Size = 40
shape.TextFrame.TextRange.Font.Bold = True
shape.TextFrame.TextRange.Font.Color.RGB = 0x0000FF  # BGR! 红色

ppt.SaveAs(r"C:\test_modified.pptx")
ppt.Close()
app.Quit()
pythoncom.CoUninitialize()
```

### C# .NET 写法

```csharp
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MsoTriState = Microsoft.Office.Core.MsoTriState;

var app = new PowerPoint.Application();
var ppt = app.Presentations.Open(@"C:\test.pptx",
    MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

var slide = ppt.Slides[1];
var shape = slide.Shapes[1];

shape.TextFrame.TextRange.Text = "新标题";
shape.TextFrame.TextRange.Font.Size = 40;
shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
shape.TextFrame.TextRange.Font.Color.RGB = 0x0000FF; // BGR! 红色

ppt.SaveAs(@"C:\test_modified.pptx");
ppt.Close();
app.Quit();
Marshal.ReleaseComObject(app);
```

### 三者对比

| 特性 | VBA | Python (pywin32) | C# (.NET Interop) |
|------|-----|-------------------|-------------------|
| 运行环境 | Office 内部 | 外部进程 | 外部进程 |
| 绑定方式 | 早绑定 | 晚绑定 (IDispatch) | 早绑定 (PIA) |
| 类型检查 | 编译时 | 运行时 | 编译时 |
| 性能 | 最快 | 较慢 | 快 |
| COM 初始化 | 自动 | 手动 CoInitialize | 自动 |
| 对象释放 | 自动 GC | Python GC | 需手动 Marshal.ReleaseComObject |
| 索引起始 | 1 | 1 | 1 |
| 颜色格式 | RGB() 函数 | BGR 整数 | BGR 整数 |
| 部署 | 嵌入文档 | pip install pywin32 | NuGet 包 |

## COM 关键差异说明

### 1. 颜色格式：BGR vs RGB

COM 内部使用 **BGR** 格式（蓝-绿-红），不是常见的 RGB！

```
颜色      RGB 值      BGR 值 (COM)
红色      #FF0000     0x0000FF
蓝色      #0000FF     0xFF0000
绿色      #00FF00     0x00FF00 (一样)
黄色      #FFD700     0x00D7FF
```

VBA 的 `RGB(255, 0, 0)` 函数内部会自动转 BGR：
```vba
RGB(255, 0, 0) = &H0000FF  ' 红色，返回的是 BGR
```

Python/C# 需要手动转：
```python
def rgb_to_bgr(r, g, b):
    return b << 16 | g << 8 | r
```

### 2. 索引：全部 1-based

COM 集合索引从 **1** 开始，不是 0：
```
Slides(1)       ← 第一页
Shapes(1)       ← 第一个元素
Paragraphs(1)   ← 第一个段落
Characters(1,3) ← 从第1个字符开始，取3个
Rows(1)         ← 第一行
Cell(1, 1)      ← 第1行第1列
```

### 3. COM 对象生命周期

| 语言 | 释放方式 |
|------|---------|
| VBA | 自动，`Set obj = Nothing` |
| Python | `pythoncom.CoInitialize()` / `CoUninitialize()` |
| C# | `Marshal.ReleaseComObject(obj)` 必须手动释放，否则 PowerPoint 进程残留 |

C# 最佳实践：
```csharp
// 反向释放（先子对象，后父对象）
Marshal.ReleaseComObject(shape);
Marshal.ReleaseComObject(slide);
Marshal.ReleaseComObject(ppt);
Marshal.ReleaseComObject(app);
GC.Collect();
GC.WaitForPendingFinalizers();
```

### 4. 线程模型

COM 使用 **STA (Single-Threaded Apartment)** 模型：

```csharp
// C# 必须标记 Main 方法
[STAThread]
static void Main(string[] args) { ... }
```

```python
# Python 必须手动初始化
pythoncom.CoInitialize()   # 进入 STA
# ... 操作 ...
pythoncom.CoUninitialize()  # 离开 STA
```

不能跨线程操作 COM 对象！

## COM 枚举常量

### PpSlideLayout (页面布局)

| 常量 | 值 | 说明 |
|------|----|------|
| ppLayoutBlank | 12 | 空白 |
| ppLayoutTitle | 1 | 标题 |
| ppLayoutText | 2 | 标题和内容 |
| ppLayoutTitleOnly | 11 | 仅标题 |

### PpEntryEffect (切换效果)

| 常量 | 值 | 说明 |
|------|----|------|
| ppEffectNone | 0 | 无 |
| ppEffectCut | 257 | 剪切 |
| ppEffectDissolve | 1537 | 溶解 |
| ppEffectFade | 1793 | 淡出 |
| ppEffectBlindsHorizontal | 769 | 水平百叶窗 |
| ppEffectBoxOut | 3073 | 盒形展开 |
| ppEffectPushDown | 3336 | 向下推入 |
| ppEffectFadeSmoothly | 3849 | 平滑淡化 |

### MsoAnimEffect (动画效果)

| 常量 | 值 | 说明 |
|------|----|------|
| msoAnimEffectAppear | 1 | 出现 |
| msoAnimEffectFly | 2 | 飞入 |
| msoAnimEffectFade | 10 | 淡入 |
| msoAnimEffectBounce | 26 | 弹跳 |
| msoAnimEffectGrowShrink | 53 | 缩放 |

### MsoAnimTriggerType (动画触发)

| 常量 | 值 | 说明 |
|------|----|------|
| msoAnimTriggerOnPageClick | 1 | 单击触发 |
| msoAnimTriggerWithPrevious | 2 | 与上一动画同时 |
| msoAnimTriggerAfterPrevious | 3 | 上一动画之后 |

### PpPlaceholderType (占位符类型)

| 常量 | 值 | 说明 |
|------|----|------|
| ppPlaceholderTitle | 1 | 标题 |
| ppPlaceholderBody | 2 | 正文 |
| ppPlaceholderCenterTitle | 3 | 居中标题 |
| ppPlaceholderSubtitle | 4 | 副标题 |
| ppPlaceholderObject | 7 | 对象 |
| ppPlaceholderTable | 8 | 表格 |
| ppPlaceholderMediaClip | 12 | 媒体 |
| ppPlaceholderSlideNumber | 13 | 页码 |
| ppPlaceholderFooter | 15 | 页脚 |

### PpSaveAsFileType (保存格式)

| 常量 | 值 | 说明 |
|------|----|------|
| ppSaveAsDefault | 11 | 默认 pptx |
| ppSaveAsPDF | 32 | PDF |
| ppSaveAsPNG | 18 | PNG |
| ppSaveAsJPG | 17 | JPG |

### PpParagraphAlignment (段落对齐)

| 常量 | 值 | 说明 |
|------|----|------|
| ppAlignLeft | 1 | 左对齐 |
| ppAlignCenter | 2 | 居中 |
| ppAlignRight | 3 | 右对齐 |
| ppAlignJustify | 4 | 两端对齐 |

## 本项目的 COM 实现

本项目提供三种 COM 实现：

| 文件 | 语言 | 说明 |
|------|------|------|
| `pptx_editor_com.py` | Python | pywin32 晚绑定，最灵活 |
| `pptx_editor_com.cs` | C# | .NET Interop 早绑定，类型安全 |

两个版本功能完全一致（32 个方法 + 自然语言意图解析），接口调用的都是同一个 `PowerPoint.Application` COM Server。

### 架构图

```
用户自然语言: "第1页标题改成红色"
          │
          ▼
   ┌─────────────┐
   │ IntentParser │   纯正则规则引擎
   │ (无 AI 依赖) │   提取: 页码 + 目标 + 操作
   └──────┬──────┘
          │ [{action, slide, target, params}]
          ▼
   ┌─────────────┐
   │ FindShape   │   三维定位: type × position × text
   └──────┬──────┘
          │ Shape 对象
          ▼
   ┌─────────────┐
   │  COM 操作    │   调用 PowerPoint.Application COM 接口
   │  (32 方法)   │   通过 IDispatch / PIA
   └──────┬──────┘
          │
          ▼
   ┌─────────────┐
   │ PowerPoint  │   COM Server (进程外)
   │    .exe      │   实际执行修改
   └─────────────┘
```

## 参考资料

- [PowerPoint VBA 对象模型](https://learn.microsoft.com/en-us/office/vba/api/overview/powerpoint/object-model)
- [PpEntryEffect 枚举](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect)
- [MsoAnimEffect 枚举](https://learn.microsoft.com/en-us/office/vba/api/office.msoanimeffect)
- [PowerPoint Interop (C#)](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.powerpoint)
