# COM 元素定位原理

## 概述

在 PowerPoint COM 自动化中，"找到正确的元素"是所有操作的前提。不同于 Web 开发中的 CSS Selector 或 XPath，COM 操作 PowerPoint 有一套独特的对象模型和定位策略。

本文档详细解析 `pptx_editor_com.py` 中的元素定位机制。

---

## 1. PowerPoint 对象模型层级

```
Application
 └── Presentations (集合)
      └── Presentation
           ├── PageSetup          → SlideWidth, SlideHeight
           ├── SlideMaster        → Theme, ThemeColorScheme
           └── Slides (集合, 1-based)
                └── Slide
                     ├── CustomLayout   → Name (版式名称)
                     ├── Layout         → 版式枚举值
                     ├── Background     → Fill
                     ├── SlideShowTransition → EntryEffect, Duration
                     ├── TimeLine       → MainSequence (动画)
                     ├── NotesPage      → Shapes.Placeholders(2)
                     ├── Comments       → Add/Delete
                     └── Shapes (集合, 1-based)
                          └── Shape
                               ├── Id            → 唯一 ID (整数)
                               ├── Name          → "Title 1", "TextBox 3"
                               ├── Type          → msoTextBox=17, msoPicture=13, msoTable=19...
                               ├── Left/Top      → 位置 (磅值 points)
                               ├── Width/Height  → 尺寸
                               ├── Rotation       → 旋转角度
                               ├── PlaceholderFormat → Type (1=标题, 2=正文, 4=副标题...)
                               ├── HasTextFrame  → Boolean
                               │    └── TextFrame
                               │         └── TextRange
                               │              ├── Text     → 纯文本
                               │              ├── Font     → Name, Size, Bold, Italic, Color.RGB
                               │              ├── Paragraphs(n) → 段落
                               │              │    └── Runs(n) → 文本运行
                               │              └── ParagraphFormat → Alignment, SpaceBefore...
                               ├── HasTable     → Boolean
                               │    └── Table
                               │         ├── Rows (集合) → Add, Delete
                               │         ├── Columns (集合) → Add, Delete
                               │         └── Cell(row, col) → Shape.TextFrame.TextRange
                               ├── Chart        → ChartData, ChartType
                               ├── Shadow       → Type, Style, Visible
                               ├── ThreeD       → RotationX/Y/Z
                               ├── Fill         → Solid, ForeColor.RGB
                               ├── Line         → ForeColor.RGB, Weight
                               └── ActionSettings → Hyperlinks
```

**关键特征：**
- 所有集合都是 **1-based 索引**（不是 0-based）
- Shape.Id 在同一 Presentation 内唯一
- Shape.Name 可能重复（如多个 "TextBox 1"）

---

## 2. Placeholder 类型体系

PowerPoint 的占位符（Placeholder）是版式（Layout）预定义的区域，有明确的类型编号：

| Type 值 | 常量名 | 中文含义 | 典型用途 |
|---------|--------|---------|---------|
| 1 | ppPlaceholderTitle | 标题 | 页面主标题 |
| 2 | ppPlaceholderBody | 正文 | 要点列表、段落 |
| 3 | ppPlaceholderCenterTitle | 居中标题 | 封面标题 |
| 4 | ppPlaceholderSubtitle | 副标题 | 封面副标题 |
| 5 | ppPlaceholderVerticalTitle | 竖排标题 | 少用 |
| 7 | ppPlaceholderVerticalBody | 竖排正文 | 少用 |
| 10 | ppPlaceholderDate | 日期 | 页脚区 |
| 11 | ppPlaceholderSlideNumber | 页码 | 页脚区 |
| 12 | ppPlaceholderFooter | 页脚 | 页脚区 |
| 15 | ppPlaceholderObject | 内容 | 可放表格/图表/SmartArt |

**注意：**
- `title` 包含 Type 1（普通标题）和 Type 3（居中标题，封面用）
- `body` 包含 Type 2（普通正文）和 Type 7（竖排正文）
- 并非所有 Shape 都是 Placeholder，访问 `PlaceholderFormat` 会抛异常

代码中的处理：

```python
try:
    pf = shape.PlaceholderFormat
    if t == "title" and pf.Type not in (1, 3): ok = False
    elif t == "subtitle" and pf.Type != 4: ok = False
    elif t == "body" and pf.Type not in (2, 7): ok = False
except:
    if t in ("title", "subtitle", "body"): ok = False
```

---

## 3. find_shape() 定位引擎

`find_shape(slide_idx, target)` 是核心定位方法，支持多维度组合查询：

### 3.1 按类型定位

```python
# 找标题（Placeholder Type 1 或 3）
p.find_shape(1, {"type": "title"})

# 找副标题（Placeholder Type 4）
p.find_shape(1, {"type": "subtitle"})

# 找正文（Placeholder Type 2 或 7）
p.find_shape(2, {"type": "body"})

# 找表格
p.find_shape(3, {"type": "table"})  # 通过 shape.HasTable 判断

# 找图片
p.find_shape(1, {"type": "picture"})  # 通过 shape.Type == 13 判断
```

### 3.2 按位置定位

将页面按 3×3 网格划分，基于 Shape 中心点判断所在区域：

```
         0%            33%           67%           100%
    ┌──────────────┬──────────────┬──────────────┐
    │              │              │              │  0%
    │    左上      │    中上      │    右上      │
    │              │              │              │  33%
    ├──────────────┼──────────────┼──────────────┤
    │              │              │              │
    │    左中      │    中中      │    右中      │
    │              │              │              │  67%
    ├──────────────┼──────────────┼──────────────┤
    │              │              │              │
    │    左下      │    中下      │    右下      │
    │              │              │              │  100%
    └──────────────┴──────────────┴──────────────┘
```

定位逻辑：

```python
cx = shape.Left + shape.Width / 2    # 中心 X
cy = shape.Top + shape.Height / 2    # 中心 Y

h = "左" if cx < sw * 0.33 else ("右" if cx > sw * 0.67 else "中")
v = "上" if cy < sh * 0.33 else ("下" if cy > sh * 0.67 else "中")

position_label = h + v  # 例如 "左上"、"中中"、"右下"
```

使用示例：

```python
# 找左上角的元素
p.find_shape(1, {"position": "左上"})

# 组合：左上角的文本框
p.find_shape(1, {"type": "textbox", "position": "左上"})
```

### 3.3 按文本内容定位

```python
# 找包含特定文本的 Shape
p.find_shape(1, {"text_match": "营收"})

# 组合：找标题中包含 "报告" 的
p.find_shape(1, {"type": "title", "text_match": "报告"})
```

### 3.4 组合定位示例

```python
# 最精确：第3页、右上角、包含"总计"的文本
p.find_shape(3, {"position": "右上", "text_match": "总计"})

# 实际场景：找到数据表格
p.find_shape(3, {"type": "table"})

# 实际场景：找到封面标题
p.find_shape(1, {"type": "title"})
```

---

## 4. inspect() 结构探查

在定位前，通常先用 `inspect()` 查看文件结构：

```python
p = PowerPointCOM(visible=False)
p.open("report.pptx")
desc = p.inspect()
p.print_structure(desc)
```

输出示例：

```
==================================================
📄 第 1 页 (Title Slide)
==================================================
  [2] Title 1 [标题] (中上) → PPT Editor Demo
       字体:Calibri Light 字号:44.0 粗:True
  [3] Subtitle 2 [副标题] (中中) → COM Automation Test Suite
       字体:Calibri 字号:20.0 粗:False

==================================================
📄 第 3 页 (Title Only)
==================================================
  [2] Title 1 [标题] (中上) → Performance Metrics
       字体:Calibri Light 字号:36.0 粗:True
  [4] Content Placeholder 4 [OBJECT] (中中) → [图片]
  [5] Table 3 (中中) → [表格 5×4]
       表格: 5×4
```

`inspect()` 返回的 JSON 结构包含：

```json
{
  "slides": [
    {
      "index": 1,
      "layout": "Title Slide",
      "elements": [
        {
          "id": 2,
          "name": "Title 1",
          "type": 14,
          "left": 100.0,
          "top": 50.0,
          "width": 600.0,
          "height": 80.0,
          "text": "PPT Editor Demo",
          "is_placeholder": true,
          "has_image": false,
          "has_chart": false,
          "has_table": false,
          "has_media": false,
          "ph_type": 3,
          "ph_type_name": "居中标题",
          "position_label": "中上",
          "paragraphs": [
            {
              "text": "PPT Editor Demo",
              "font": "Calibri Light",
              "size": 44.0,
              "bold": true,
              "color": 0
            }
          ]
        }
      ]
    }
  ]
}
```

当元素没有文本但包含图片、图表、表格或媒体时，`inspect()` 会通过 `has_image`、`has_chart`、`has_table`、`has_media` 明确标记，`print_structure()` 也会优先显示这些标签而不是 `(无)`。

---

## 5. LLM 意图解析中的定位

`pptx_editor_llm.py` 将 inspect() 的结构作为上下文传给 LLM，LLM 返回 JSON 操作指令：

```json
[
  {
    "action": "modify_text",
    "slide": 1,
    "target": {"type": "title"},
    "text": "新标题"
  },
  {
    "action": "modify_font",
    "slide": 1,
    "target": {"type": "title"},
    "font_size": 48,
    "bold": true,
    "color": "FF0000"
  }
]
```

### 3.4 按名称定位 (name)

每个 Shape 都有 `.Name` 属性（如 "TextBox 3"、"Title 1"），可用于精确定位：

```python
target = {"name": "TextBox 3"}
# shape.Name.lower() 包含 target["name"].lower() 即匹配
```

### 3.5 按索引定位 (index)

Shapes 集合的遍历顺序是固定的（1-based），可直接用索引定位：

```python
target = {"index": 3}  # 第3个形状
```

适用于密集页面中 type + position 无法区分多个同类元素的场景。

**定位策略优先级（LLM 应遵循）：**

1. **Placeholder 类型** — 最可靠：`{"type": "title"}`, `{"type": "body"}`
2. **文本匹配** — 内容已知时：`{"text_match": "营收数据"}`
3. **位置区域** — 布局已知时：`{"position": "左上"}`
4. **名称匹配** — 已知形状名时：`{"name": "TextBox 3"}`
5. **索引定位** — 最后手段：`{"index": 3}`
6. **组合查询** — 多条件交叉：`{"type": "body", "position": "左中", "name": "Content"}`

---

## 6. Shape.Type 枚举值

COM 中每个 Shape 都有 `.Type` 属性，常见值：

| 值 | 常量 | 含义 |
|----|------|------|
| 1 | msoAutoShape | 自选图形（矩形、圆等） |
| 5 | msoFreeform | 自由曲线 |
| 6 | msoGroup | 组合形状 |
| 11 | msoLinkedPicture | 链接图片 |
| 13 | msoPicture | 嵌入图片 |
| 14 | msoPlaceholder | 占位符 |
| 15 | msoScriptAnchor | 脚本 |
| 17 | msoTextBox | 文本框 |
| 19 | msoTable | 表格 |
| 16 | msoMedia | 媒体（音视频） |
| 24 | msoSmartArt | SmartArt |
| 3 | msoChart | 图表 |

**注意：** Placeholder（Type=14）是一个容器类型，实际内容可能是文本、表格或图片。判断 Placeholder 的用途需要看 `PlaceholderFormat.Type`。

---

## 7. 颜色系统：BGR 而非 RGB

COM 中的 `Color.RGB` 属性使用 **BGR 格式**：

```
0xBBGGRR

红色: 0x0000FF  (不是 0xFF0000)
绿色: 0x00FF00  (一样)
蓝色: 0xFF0000  (不是 0x0000FF)
白色: 0xFFFFFF
黑色: 0x000000
```

转换公式：

```python
# RGB → BGR
def rgb_to_bgr(r, g, b):
    return (b << 16) | (g << 8) | r

# BGR → RGB
def bgr_to_rgb(bgr):
    r = bgr & 0xFF
    g = (bgr >> 8) & 0xFF
    b = (bgr >> 16) & 0xFF
    return r, g, b
```

---

## 8. 坐标系统

COM 中的 Left/Top/Width/Height 单位是 **磅 (points)**，1 inch = 72 points。

```
标准 16:9 幻灯片: 960 × 540 points (13.33 × 7.5 inches)
标准  4:3 幻灯片: 720 × 540 points (10 × 7.5 inches)
```

常用换算：

```python
from pptx.util import Inches, Pt, Emu

# 1 inch = 72 points = 914400 EMU
# COM 返回的是 points
# python-pptx 返回的是 EMU
```

---

## 9. 定位踩坑记录

### 坑1：Placeholder 访问异常
非占位符 Shape 访问 `PlaceholderFormat` 会抛 `com_error`，必须 try-except。

### 坑2：Title 有两种类型
封面页用 Type=3（CenterTitle），内容页用 Type=1（Title）。find_shape 里要同时匹配。

### 坑3：Body vs Subtitle
封面页的第二个占位符是 Subtitle（Type=4），不是 Body（Type=2）。很多人搞混。

### 坑4：Shape.Name 不唯一
同一页可以有多个 "TextBox 1"，不要依赖 Name 做定位，用类型+位置组合更可靠。

### 坑5：集合是 1-based
`Slides(1)` 是第一页，`Table.Cell(1,1)` 是左上角。传 0 会报 "out of range"。

### 坑6：Shape 顺序不稳定
`Shapes` 集合的顺序取决于添加顺序和 ZOrder，不保证"从上到下"或"从左到右"。用 position_label 更可靠。
