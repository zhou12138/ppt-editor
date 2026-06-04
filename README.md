# PPT Editor — PPTX 自然语言编辑器

用自然语言指令编辑 PowerPoint 文件。说人话，改 PPT。

## 双版本

| 版本 | 文件 | 依赖 | 平台 |
|------|------|------|------|
| python-pptx 版 | `pptx_editor.py` | python-pptx | 全平台 |
| COM 版 | `pptx_editor_com.py` | pywin32 + Office | Windows |

## 安装

```bash
# python-pptx 版（全平台）
pip install python-pptx

# COM 版（Windows Only）
pip install pywin32
# + 需要安装 Microsoft Office (PowerPoint)
```

## 用法

```bash
# 单条指令
python pptx_editor_com.py report.pptx "把标题改成「Q3 总结」"
python pptx_editor_com.py report.pptx "把标题改成「Q3 总结」" --output result.pptx

# 查看结构
python pptx_editor_com.py report.pptx --inspect

# 交互模式
python pptx_editor_com.py report.pptx --interactive

# 导出所有页为图片
python pptx_editor_com.py report.pptx --export-images
```

## Windows 本地测试

```bash
demo_test.bat          # 运行所有测试
demo_test.bat clean    # 一键清理测试输出
```

---

## COM 版详细文档

### 整体架构

```
用户输入自然语言指令
       │
       ▼
 parse_intent()          ← 纯正则规则引擎，不依赖 AI
       │
       ▼
 [{action, slide, target, params}, ...]   ← 意图列表
       │
       ▼
 find_shape()            ← 按 type/position/text/name/index 五维定位元素
       │
       ▼
 modify / delete / animate / transition   ← 执行操作
       │
       ▼
 save()                  ← 保存修改后的文件
```

### 意图解析详解 (`parse_intent`)

意图解析是纯正则规则引擎，从自然语言中提取三要素：**页码 + 目标元素 + 操作**。

#### 1. 页码提取

```python
re.search(r'第(\d+)[页张]', instruction)
```

- `"第2页标题改成xxx"` → `slide_num = 2`
- 不指定页码 → 对所有页执行

#### 2. 目标元素定位 (`target`)

目标由三个维度组合定位：

| 维度 | 关键词 | 匹配规则 |
|------|--------|----------|
| **类型** | 标题/title/题目 → `title` | 占位符类型 1 或 3 |
| | 副标题/subtitle → `subtitle` | 占位符类型 4 |
| | 正文/内容/body → `body` | 占位符类型 2 或 7 |
| | 表格 → `table` | `shape.HasTable` |
| | 图片/picture → `picture` | `shape.Type == 13` |
| **位置** | 左上/右上/左下/右下/居中 | 元素中心点在页面九宫格的位置 |
| **文本** | `"引号中的文字"` | `shape.TextFrame.TextRange.Text` 包含该文字 |
| **名称** | `name:XXX` 或 `名称:XXX` | `shape.Name` 包含该子串（不区分大小写） |
| **索引** | `#N` 或 `第N个` | Shapes 集合中第 N 个元素（1-based） |

位置计算逻辑：
```
页面宽度 sw, 高度 sh
元素中心 (cx, cy) = (left + width/2, top + height/2)
水平: cx < sw*0.33 → 左, cx > sw*0.67 → 右, 否则 → 中
垂直: cy < sh*0.33 → 上, cy > sh*0.67 → 下, 否则 → 中
```

#### 3. 操作识别

| 操作 | 触发正则/关键词 | 提取的参数 |
|------|----------------|-----------|
| **修改文本** | `改/换/替换/变 + 成/为 + "新文本"` | `new_text` |
| **修改字号** | `字号 + 改/调/设为 + 数字` | `font_size` |
| **放大字号** | `大一点` | `font_size_factor: 1.3` |
| **缩小字号** | `小一点` | `font_size_factor: 0.75` |
| **加粗** | `加粗/粗体/bold` | `bold: True` |
| **改颜色** | `改/换/变/调 + 成 + 颜色名` | `color: BGR值` |
| **改字体** | `字体 + 改/换/用 + 字体名` | `font_name` |
| **删除** | `删除/删掉/去掉` | — |
| **添加动画** | `添加/加/设置 + 动画 + 效果名` | `effect` |
| **切换效果** | `切换 + 效果 + 效果名` | `transition` |
| **导出PDF** | `导出pdf/导出PDF/转pdf/转PDF` | — |

#### 4. 引号歧义消解

引号内容可能是**定位条件**（找到包含该文字的元素）或**操作参数**（改成该文字）：

```python
# 如果引号前面是 "改/换/替换/变+成/为"，则为操作参数
re.search(r'(?:改|换|替换|变)[成为]?\s*$', before_quote)

# 否则为定位条件
target["text_match"] = quoted_text
```

示例：
- `把"旧标题"改成"新标题"` → 定位: text_match="旧标题"，操作: new_text="新标题"
- `把标题改成"新标题"` → 定位: type=title，操作: new_text="新标题"

#### 5. 支持的颜色

| 中文 | 英文 | BGR 值 |
|------|------|--------|
| 红/红色 | red | `0x0000FF` |
| 蓝/蓝色 | blue | `0xFF0000` |
| 绿/绿色 | green | `0x00AA00` |
| 黄/黄色 | — | `0x00D7FF` |
| 黑/黑色 | black | `0x000000` |
| 白/白色 | white | `0xFFFFFF` |
| 灰/灰色 | — | `0x888888` |
| 橙/橙色 | — | `0x008CFF` |
| 紫/紫色 | — | `0x800080` |
| 粉/粉色 | — | `0xB469FF` |

> ⚠️ COM 使用 BGR 颜色格式，不是 RGB！红色 `#FF0000` 在 COM 中是 `0x0000FF`。

#### 6. 动画效果映射

| 中文 | 英文 | MsoAnimEffect 值 |
|------|------|------------------|
| 出现 | appear | 1 |
| 飞入 | fly | 2 |
| 淡入 | fade | 10 |
| 缩放 | zoom | 53 |
| 弹跳 | bounce | 26 |

#### 7. 切换效果映射

| 中文 | 英文 | PpEntryEffect 值 |
|------|------|------------------|
| 淡化 | fade | 3849 (ppEffectFadeSmoothly) |
| 推入 | push | 3336 |
| 擦除 | wipe | 769 (ppEffectBlindsHorizontal) |
| 分裂 | split | 3073 (ppEffectBoxOut) |
| 溶解 | dissolve | 1537 (ppEffectDissolve) |
| 剪切 | cut | 257 (ppEffectCut) |
| 无 | none | 0 (ppEffectNone) |

---

### COM 接口调用清单

以下是 `pptx_editor_com.py` 使用的所有 Windows COM (PowerPoint.Application) 接口：

#### 应用生命周期

| COM 调用 | 用途 | 代码位置 |
|----------|------|----------|
| `pythoncom.CoInitialize()` | 初始化 COM 运行时 | `__init__` |
| `win32com.client.Dispatch("PowerPoint.Application")` | 启动 PowerPoint 进程 | `__init__` |
| `app.Visible = True/False` | 控制 PowerPoint 窗口可见性 | `__init__` |
| `app.Quit()` | 关闭 PowerPoint 进程 | `close` |
| `pythoncom.CoUninitialize()` | 释放 COM 运行时 | `close` |

#### 文件操作

| COM 调用 | 用途 |
|----------|------|
| `app.Presentations.Open(path, False, False, False)` | 打开 PPTX 文件（只读=否, 不含标题=否, 不含窗口=否） |
| `prs.Save()` | 保存到原文件 |
| `prs.SaveAs(path)` | 另存为 PPTX |
| `prs.SaveAs(path, 32)` | 另存为 PDF（ppSaveAsPDF=32） |
| `prs.Close()` | 关闭演示文稿 |

#### 页面信息

| COM 调用 | 用途 |
|----------|------|
| `prs.Slides.Count` | 总页数 |
| `prs.Slides(index)` | 按索引获取页面（1-based） |
| `prs.PageSetup.SlideWidth` | 页面宽度（用于位置计算） |
| `prs.PageSetup.SlideHeight` | 页面高度 |
| `slide.CustomLayout.Name` | 页面布局名称 |
| `slide.Layout` | 页面布局类型枚举 |

#### 元素 (Shape) 操作

| COM 调用 | 用途 |
|----------|------|
| `slide.Shapes` | 遍历页面上所有元素 |
| `shape.Id` | 元素 ID |
| `shape.Name` | 元素名称 |
| `shape.Type` | 元素类型（13=图片等） |
| `shape.Left / Top / Width / Height` | 位置和尺寸 |
| `shape.Delete()` | 删除元素 |

#### 占位符

| COM 调用 | 用途 |
|----------|------|
| `shape.PlaceholderFormat` | 获取占位符信息 |
| `pf.Type` | 占位符类型（1=TITLE, 2=BODY, 3=CENTER_TITLE, 4=SUBTITLE, 7=OBJECT 等） |

#### 文本操作

| COM 调用 | 用途 |
|----------|------|
| `shape.HasTextFrame` | 是否有文本框 |
| `shape.TextFrame.TextRange.Text` | 读/写文本内容 |
| `textRange.Paragraphs()` | 获取段落集合 |
| `textRange.Paragraphs(index)` | 按索引获取段落（1-based） |
| `paragraph.Text` | 段落文本 |

#### 字体样式

| COM 调用 | 用途 |
|----------|------|
| `textRange.Font.Name` | 读/写字体名称 |
| `textRange.Font.Size` | 读/写字号 |
| `textRange.Font.Bold` | 读/写粗体 |
| `textRange.Font.Italic` | 读/写斜体 |
| `textRange.Font.Color.RGB` | 读/写字体颜色（BGR 格式） |

#### 表格

| COM 调用 | 用途 |
|----------|------|
| `shape.HasTable` | 是否是表格 |
| `shape.Table` | 获取表格对象 |
| `table.Rows.Count` | 行数 |
| `table.Columns.Count` | 列数 |
| `table.Cell(row, col).Shape.TextFrame.TextRange.Text` | 读取单元格文本 |

#### 动画 (COM 独有)

| COM 调用 | 用途 |
|----------|------|
| `slide.TimeLine.MainSequence.AddEffect(Shape, effectId, trigger)` | 添加动画效果 |
| 参数 `trigger=1` | 点击触发（msoAnimTriggerOnPageClick） |

#### 切换效果 (COM 独有)

| COM 调用 | 用途 |
|----------|------|
| `slide.SlideShowTransition.EntryEffect` | 设置页面切换效果 |
| `slide.SlideShowTransition.Duration` | 设置切换持续时间（秒） |

#### 导出 (COM 独有)

| COM 调用 | 用途 |
|----------|------|
| `prs.SaveAs(path, 32)` | 导出 PDF（ppSaveAsPDF=32） |
| `slide.Export(path, "PNG", width, height)` | 导出单页为 PNG 图片 |

---

### 已覆盖的编辑场景

| 分类 | 场景 | 状态 |
|------|------|------|
| **文本** | 修改文本内容 | ✅ |
| **文本** | 删除元素 | ✅ |
| **样式** | 修改字号（绝对值） | ✅ |
| **样式** | 修改字号（相对缩放） | ✅ |
| **样式** | 加粗/取消加粗 | ✅ |
| **样式** | 斜体 | ✅（代码支持，意图解析未接入） |
| **样式** | 修改字体颜色 | ✅ |
| **样式** | 修改字体名称 | ✅ |
| **动画** | 添加入场动画 | ✅ COM 独有 |
| **切换** | 设置页面切换效果 | ✅ COM 独有 |
| **导出** | 导出 PDF | ✅ COM 独有 |
| **导出** | 导出 PNG 图片 | ✅ COM 独有 |
| **查看** | 查看 PPT 结构 | ✅ |
| **查看** | 导出结构 JSON | ✅ |

### 未覆盖的场景（待实现）

| 分类 | 场景 | 难度 |
|------|------|------|
| **文本** | 添加新文本框 | 低 |
| **文本** | 修改单个段落/部分文字 | 中 |
| **样式** | 下划线/删除线 | 低 |
| **样式** | 文本对齐（居左/居中/居右） | 低 |
| **样式** | 修改元素背景/填充色 | 中 |
| **样式** | 修改元素边框 | 中 |
| **布局** | 移动元素位置 | 中 |
| **布局** | 调整元素大小 | 中 |
| **布局** | 添加/删除页面 | 低 |
| **布局** | 调整页面顺序 | 低 |
| **表格** | 修改单元格内容 | 中 |
| **表格** | 添加/删除行列 | 中 |
| **图片** | 插入图片 | 中 |
| **图片** | 替换图片 | 中 |
| **动画** | 删除/修改已有动画 | 中 |
| **动画** | 设置动画顺序和延迟 | 中 |
| **母版** | 修改母版/布局 | 高 |
| **高级** | 跨轮次上下文记忆 | 高 |
| **高级** | 歧义消解（多匹配时询问） | 高 |
| **高级** | 批量指令（一句话多个操作） | 中（已部分支持） |

---

### COM vs python-pptx 能力对比

| 能力 | python-pptx | COM |
|------|:-----------:|:---:|
| 跨平台 | ✅ | ❌ Windows Only |
| 无需 Office | ✅ | ❌ 需要 |
| 修改文本/样式 | ✅ | ✅ |
| 添加动画 | ❌ | ✅ |
| 切换效果 | ❌ | ✅ |
| 导出 PDF | ❌ | ✅ |
| 导出图片 | ❌ | ✅ |
| 读取渲染后样式 | ❌ | ✅ |
| 无头运行 | ✅ | ✅（Visible=False） |
| 并发安全 | ✅ | ❌ COM 单线程 |

> COM 版通过 `pythoncom.CoInitialize()` 初始化 COM 单线程套间（STA），不支持多线程并发操作同一个 PowerPoint 实例。
