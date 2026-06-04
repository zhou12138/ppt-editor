# CLAUDE.md — PPT Editor 项目指南

## 项目结构

```
pptx_editor_com.py          # COM 引擎（所有 PowerPoint 操作方法）
pptx_editor_llm.py          # 统一入口（三种执行模式）
pptx_editor.py              # 纯规则引擎（python-pptx，跨平台）
COM_STANDARD.md             # COM 接口标准文档（对象模型/枚举常量/VBA 对照）
demo_test_skill.bat         # skill 脚本 smoke test（含 clean）
skills/                     # 三种部署场景的 skill 文档
```

## 编辑 PPTX 的方式

你（Claude Code）在本地 Windows 上编辑 PPTX，推荐两种方式：

### 方式 B：直接生成 Python 脚本（推荐）

1. 先 inspect 查看结构：
```bash
python pptx_editor_llm.py deck.pptx --inspect
```

2. 生成 edit.py 脚本，用 `ppt`（PowerPointCOM 实例）和 `filepath` 操作：
```python
# edit.py — ppt 和 filepath 已注入为全局变量
structure = ppt.inspect()
shapes = ppt.find_shape(1, {"type": "title"})
for s in shapes:
    ppt.modify_font(s, bold=True, color=0x0000FF)
```

3. 执行：
```bash
python pptx_editor_llm.py deck.pptx --exec-script edit.py --output result.pptx
```

### 方式 C：生成 JSON Actions

```bash
python pptx_editor_llm.py deck.pptx --exec-actions '[{"action":"modify_font","slide":1,"target":{"type":"title"},"params":{"bold":true}}]'
python pptx_editor_llm.py deck.pptx --exec-actions actions.json --output result.pptx
```

## 关键注意事项

- **BGR 颜色！** COM 用 BGR 不是 RGB。红=0x0000FF，蓝=0xFF0000。公式：`BGR = R + G*256 + B*65536`
- **1-Based 索引** — 所有 COM 索引从 1 开始
- **不需要 API Key** — 方式 B/C 完全本地执行，不调任何 LLM API
- **COM 生命周期** — 脚本框架已处理 CoInitialize/Close，你的 edit.py 里不需要管

## PowerPointCOM 常用方法速查

### 结构查看
- `ppt.inspect()` → 返回完整结构 dict
- `ppt.find_shape(slide, target)` → target: `{"type":"title/subtitle/body/table/picture"}`, `{"text_match":"xxx"}`, `{"position":"左上/中中/右下"}`

### 文本
- `ppt.modify_text(shape, "新文本")`
- `ppt.modify_font(shape, font_size=24, bold=True, italic=False, color=0x0000FF, font_name="微软雅黑", font_size_factor=1.5)`
- `ppt.set_alignment(shape, "center")` — left/center/right/justify

### 形状外观
- `ppt.set_fill(shape, color_bgr)`
- `ppt.set_border(shape, color_bgr=0xFF0000, weight=2)`
- `ppt.set_shadow(shape, shadow_type)`
- `ppt.set_3d_rotation(shape, x, y, z)`

### 位置/大小（单位 points，72pt = 1 inch）
- `ppt.move_shape(shape, left=100, top=200)`
- `ppt.resize_shape(shape, width=400, height=300)`
- `ppt.delete_shape(shape)`

### 添加元素
- `ppt.add_textbox(slide, "内容", left, top, width, height)`
- `ppt.add_picture(slide, "image.png", left, top, width, height)`
- `ppt.add_chart(slide, chart_type, data, left, top, width, height)`
- `ppt.add_freeform(slide, points_list)`

### 幻灯片管理
- `ppt.add_slide(index, layout)` — layout: 1=标题, 2=标题+正文, 7=空白, 12=空白
- `ppt.delete_slide(slide_idx)`
- `ppt.move_slide(slide_idx, new_pos)`
- `ppt.duplicate_slide(slide_idx)`

### 表格
- `ppt.modify_cell(slide, target, row, col, "text")` — 1-based
- `ppt.add_table_row(shape)` / `ppt.delete_table_row(shape, row)`
- `ppt.add_table_column(shape)` / `ppt.delete_table_column(shape, col)`

### COM 独有
- `ppt.add_animation(slide, shape, "fade")` — appear/fly/fade/zoom/bounce
- `ppt.remove_animation(slide, anim_index)`
- `ppt.set_transition(slide, "fade", duration)` — fade/push/wipe/split/dissolve
- `ppt.export_pdf("output.pdf")`
- `ppt.export_image(slide, "slide.png", width, height)`

### 其他
- `ppt.set_notes(slide, "备注文本")`
- `ppt.add_comment(slide, "评论", "作者", x, y)`
- `ppt.add_hyperlink(shape, "https://...")`
- `ppt.set_slide_background(slide, color_bgr)`
- `ppt.set_line_spacing(shape, factor)`
- `ppt.merge_presentations(["other.pptx"], "merged.pptx")`

## JSON Actions 格式（方式 C）

```json
[
  {"action": "modify_font", "slide": 1, "target": {"type": "title"}, "params": {"bold": true, "color": 255}},
  {"action": "add_textbox", "slide": 1, "params": {"text": "Hello", "left": 100, "top": 100, "width": 300, "height": 50}},
  {"action": "transition", "slide": 1, "params": {"transition": "fade", "duration": 1.5}}
]
```

完整 action 列表见 `pptx_editor_llm.py` 中的 `SYSTEM_PROMPT`。
