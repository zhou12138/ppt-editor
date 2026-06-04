# PyWin32 COM 自动化编辑 PowerPoint 技术文档

## 1. 架构概述：python-pptx vs pywin32 COM

| 特性 | python-pptx | pywin32 COM |
|------|------------|-------------|
| 平台 | 跨平台（Linux/Mac/Windows） | 仅 Windows |
| 依赖 | 纯 Python，无需 Office | 需安装 Microsoft Office |
| 原理 | 直接读写 .pptx XML | 通过 COM 接口驱动 PowerPoint 进程 |
| 功能覆盖 | 基础（文本、形状、图片、表格） | 完整 PowerPoint API（动画、3D、图表、母版等） |
| 渲染能力 | 无（只改数据） | 完整（可导出 PDF/图片） |
| 适用场景 | 批量生成模板化 PPT | 精细编辑、格式调整、导出 |

**本项目选择 pywin32 COM**，因为需要完整控制 PowerPoint 的所有功能，包括动画、3D 效果、图表样式等 python-pptx 无法覆盖的能力。

---

## 2. COM 工作原理

### 生命周期

```
CoInitialize()          # 初始化 COM 线程
    ↓
Dispatch("PowerPoint.Application")  # 启动/连接 PowerPoint 进程
    ↓
app.Presentations.Open(path)         # 打开文件
    ↓
操作 Slides / Shapes / TextFrame ... # 编辑内容
    ↓
prs.SaveAs(path, format)            # 保存文件
    ↓
prs.Close()                          # 关闭演示文稿
app.Quit()                           # 退出 PowerPoint
    ↓
CoUninitialize()                     # 释放 COM 线程
```

### 核心代码模式

```python
import win32com.client
import pythoncom

pythoncom.CoInitialize()
try:
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = 1  # 必须可见，否则某些操作失败
    prs = app.Presentations.Open(r"C:\path\to\file.pptx")
    
    # 操作...
    slide = prs.Slides(1)
    shape = slide.Shapes(1)
    shape.TextFrame.TextRange.Text = "新标题"
    
    prs.Save()
    prs.Close()
    app.Quit()
finally:
    pythoncom.CoUninitialize()
```

> **注意**：COM 对象索引从 **1** 开始，不是 0。`Slides(1)` 是第一张幻灯片。

---

## 3. COM 对象模型

```
Application
 └─ Presentations (集合)
     └─ Presentation
         ├─ Slides (集合)
         │   └─ Slide
         │       ├─ Shapes (集合)
         │       │   └─ Shape
         │       │       ├─ TextFrame
         │       │       │   └─ TextRange
         │       │       │       ├─ Text (文本内容)
         │       │       │       ├─ Font (字体、颜色、大小)
         │       │       │       └─ ParagraphFormat (对齐、行距)
         │       │       ├─ Table
         │       │       │   └─ Cell(row, col)
         │       │       │       └─ Shape → TextFrame
         │       │       ├─ Chart
         │       │       │   ├─ ChartData
         │       │       │   └─ SeriesCollection
         │       │       ├─ Fill (填充)
         │       │       ├─ Line (边框)
         │       │       ├─ ThreeD (3D 效果)
         │       │       └─ AnimationSettings
         │       ├─ SlideShowTransition (切换效果)
         │       └─ Background
         ├─ SlideMaster
         └─ PageSetup
```

---

## 4. 颜色陷阱：BGR 而非 RGB

**这是最常见的坑。** COM 接口使用 BGR 格式（`0xBBGGRR`），而非常规 RGB。

```python
# ❌ 错误：想要红色 (255,0,0)，直接传 0xFF0000
shape.TextFrame.TextRange.Font.Color.RGB = 0xFF0000  # 实际显示蓝色！

# ✅ 正确：红色应该是 0x0000FF
shape.TextFrame.TextRange.Font.Color.RGB = 0x0000FF  # 红色

# 转换函数
def rgb_to_bgr(r, g, b):
    return b << 16 | g << 8 | r

red = rgb_to_bgr(255, 0, 0)    # → 0x0000FF
blue = rgb_to_bgr(0, 0, 255)   # → 0xFF0000
white = rgb_to_bgr(255,255,255) # → 0xFFFFFF
```

在 `pptx_editor_com.py` 中，所有颜色参数均接受 RGB 元组，内部自动转换为 BGR。

---

## 5. pptx_editor_com.py 方法分类（70+ 方法）

### 📝 文本操作
| 方法 | 说明 |
|------|------|
| `set_text(slide, shape, text)` | 设置形状文本 |
| `append_text(slide, shape, text)` | 追加文本 |
| `set_text_by_name(slide, name, text)` | 按形状名称设置文本 |
| `replace_text(old, new)` | 全局替换文本 |
| `set_notes(slide, text)` | 设置备注 |

### 🔤 字体与格式
| 方法 | 说明 |
|------|------|
| `set_font(slide, shape, name, size, bold, italic, color)` | 设置字体属性 |
| `set_font_color(slide, shape, r, g, b)` | 设置字体颜色（RGB 输入，自动转 BGR） |
| `set_paragraph_alignment(slide, shape, align)` | 段落对齐 |
| `set_line_spacing(slide, shape, spacing)` | 行距 |
| `set_text_shadow(slide, shape)` | 文字阴影 |

### 🔷 形状操作
| 方法 | 说明 |
|------|------|
| `add_shape(slide, type, left, top, width, height)` | 添加形状 |
| `set_shape_fill(slide, shape, r, g, b)` | 形状填充色 |
| `set_shape_border(slide, shape, r, g, b, weight)` | 形状边框 |
| `move_shape(slide, shape, left, top)` | 移动形状 |
| `resize_shape(slide, shape, width, height)` | 调整大小 |
| `delete_shape(slide, shape)` | 删除形状 |
| `duplicate_shape(slide, shape)` | 复制形状 |
| `set_shape_rotation(slide, shape, angle)` | 旋转 |
| `group_shapes(slide, shape_indices)` | 组合形状 |

### 🖼️ 图片
| 方法 | 说明 |
|------|------|
| `add_image(slide, path, left, top, width, height)` | 插入图片 |
| `replace_image(slide, shape, path)` | 替换图片 |

### 📊 表格
| 方法 | 说明 |
|------|------|
| `add_table(slide, rows, cols, left, top, width, height)` | 添加表格 |
| `set_cell_text(slide, shape, row, col, text)` | 设置单元格文本 |
| `set_cell_fill(slide, shape, row, col, r, g, b)` | 单元格填充色 |
| `merge_cells(slide, shape, r1, c1, r2, c2)` | 合并单元格 |
| `set_table_border(slide, shape, weight, r, g, b)` | 表格边框 |

### 📈 图表
| 方法 | 说明 |
|------|------|
| `add_chart(slide, type, data, left, top, w, h)` | 添加图表 |
| `modify_chart_data(slide, shape, data)` | 修改图表数据 |
| `set_chart_title(slide, shape, title)` | 图表标题 |
| `set_chart_style(slide, shape, style)` | 图表样式 |

### 🎬 动画与切换
| 方法 | 说明 |
|------|------|
| `add_animation(slide, shape, effect_type)` | 添加动画效果 |
| `set_animation_timing(slide, shape, duration, delay)` | 动画时间 |
| `set_slide_transition(slide, type, duration)` | 幻灯片切换 |
| `set_transition_sound(slide, sound_path)` | 切换声音 |

### 🎲 3D 效果
| 方法 | 说明 |
|------|------|
| `set_3d_rotation(slide, shape, x, y, z)` | 3D 旋转 |
| `set_3d_depth(slide, shape, depth)` | 3D 深度 |
| `set_3d_bevel(slide, shape, type, w, h)` | 棱台效果 |
| `set_3d_material(slide, shape, material)` | 材质 |
| `set_3d_lighting(slide, shape, type)` | 光照 |

### 📄 幻灯片管理
| 方法 | 说明 |
|------|------|
| `add_slide(layout_index)` | 添加幻灯片 |
| `delete_slide(index)` | 删除幻灯片 |
| `duplicate_slide(index)` | 复制幻灯片 |
| `move_slide(from, to)` | 移动幻灯片 |
| `set_background(slide, r, g, b)` | 设置背景色 |
| `set_background_image(slide, path)` | 背景图片 |

### 📤 导出
| 方法 | 说明 |
|------|------|
| `export_pdf(output_path)` | 导出 PDF |
| `export_images(output_dir, format)` | 导出为图片 |
| `export_slide_image(slide, path, w, h)` | 导出单张幻灯片 |

### 🔍 查询
| 方法 | 说明 |
|------|------|
| `get_slide_count()` | 获取幻灯片数量 |
| `get_shape_list(slide)` | 获取形状列表 |
| `get_text(slide, shape)` | 获取文本内容 |
| `get_shape_properties(slide, shape)` | 获取形状属性 |

---

## 6. demo_test.py —— 活的测试套件

`demo_test.py` 是一个完整的集成测试，运行后会生成一个精美的 **6 页演示文稿**，覆盖所有主要功能：

| 幻灯片 | 内容 | 测试的功能 |
|--------|------|-----------|
| 第 1 页 | 封面 | 标题文本、字体设置、背景色 |
| 第 2 页 | 文本排版 | 多段落、对齐方式、行距、文字颜色 |
| 第 3 页 | 形状与图片 | 各种形状、填充、边框、图片插入 |
| 第 4 页 | 表格 | 表格创建、单元格样式、合并单元格 |
| 第 5 页 | 图表 | 柱状图/饼图、数据修改、图表样式 |
| 第 6 页 | 动画与 3D | 动画效果、3D 旋转、切换效果 |

### 运行方式

```bash
cd pptx-editor
python demo_test.py
# 生成 demo_output.pptx
```

每个功能区域有独立的 `test_xxx()` 函数，可单独调用进行调试。

---

## 7. Session 0 问题（无头服务器的 COM 陷阱）

### 问题描述

Windows 服务（包括远程桌面断开后的后台进程）运行在 **Session 0**，这是一个没有桌面的隔离会话。PowerPoint COM 在 Session 0 中会出现：

- `Presentations.Open()` 挂起或报错
- `SaveAs()` 弹出不可见的对话框导致死锁
- 图表操作失败（需要 GDI 渲染）
- 导出 PDF/图片失败

### 原因

PowerPoint 是一个 GUI 应用程序，很多 COM 方法内部依赖窗口消息循环和 GDI 渲染。Session 0 没有交互式桌面，这些操作无法完成。

### 解决方案

| 方案 | 说明 | 推荐度 |
|------|------|--------|
| **保持 RDP 登录** | 始终保持一个远程桌面会话连接 | ⭐⭐⭐ 最简单 |
| **使用 `tscon` 重定向** | 断开 RDP 时将会话保持在桌面 | ⭐⭐⭐ 推荐 |
| **计划任务 + 交互模式** | 用任务计划程序在交互会话中运行 | ⭐⭐ |
| **自动登录 + 锁屏** | 设置 Windows 自动登录，COM 进程在该会话中运行 | ⭐⭐ |
| **VBS 包装** | 用 VBScript 创建 Shell 对象启动 | ⭐ |

### tscon 命令示例（推荐）

断开 RDP 时执行以下命令，保持桌面会话：

```bat
:: 将当前 RDP 会话重定向到控制台，保持桌面可用
tscon %sessionname% /dest:console
```

### 检测当前 Session

```python
import ctypes
def get_session_id():
    pid = ctypes.windll.kernel32.GetCurrentProcessId()
    session_id = ctypes.c_ulong()
    ctypes.windll.kernel32.ProcessIdToSessionId(pid, ctypes.byref(session_id))
    return session_id.value

sid = get_session_id()
if sid == 0:
    print("⚠️ 当前在 Session 0，COM 操作可能失败！")
else:
    print(f"✅ 当前在 Session {sid}，COM 操作正常。")
```

---

## 附录：常用 PowerPoint 常量

```python
# 形状类型 (MsoAutoShapeType)
msoShapeRectangle = 1
msoShapeRoundedRectangle = 5
msoShapeOval = 9

# 段落对齐 (PpParagraphAlignment)
ppAlignLeft = 1
ppAlignCenter = 2
ppAlignRight = 3

# 动画效果 (MsoAnimEffect)
msoAnimEffectAppear = 1
msoAnimEffectFly = 2
msoAnimEffectFade = 10

# 图表类型 (XlChartType)
xlColumnClustered = 51
xlPie = 5
xlLine = 4

# 导出格式 (PpSaveAsFileType)
ppSaveAsDefault = 11      # .pptx
ppSaveAsPDF = 32           # .pdf
ppSaveAsJPG = 17           # .jpg
ppSaveAsPNG = 18           # .png
```
