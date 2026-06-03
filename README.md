# PPT Editor — PPTX 自然语言编辑器

用自然语言指令编辑 PowerPoint 文件。说人话，改 PPT。

## 双版本

| 版本 | 文件 | 依赖 | 平台 |
|------|------|------|------|
| python-pptx 版 | `pptx_editor.py` | python-pptx | 全平台 |
| COM 版 | `pptx_editor_com.py` | pywin32 + Office | Windows |

### COM 版独有能力

- 🎬 添加动画（淡入/飞入/缩放/弹跳）
- 🔄 页面切换效果（淡化/推入/擦除）
- 📄 导出 PDF
- 🖼️ 导出图片（PNG）

## 安装

```bash
# python-pptx 版（全平台）
pip install python-pptx

# COM 版（Windows Only）
pip install pywin32
# + 需要安装 Microsoft Office
```

## 用法

### 单条指令

```bash
# python-pptx 版
python pptx_editor.py report.pptx "把标题改成「Q3 季度总结」"
python pptx_editor.py report.pptx "第2页正文字号改成18"

# COM 版
python pptx_editor_com.py report.pptx "给第1页标题添加动画淡入"
python pptx_editor_com.py report.pptx "导出PDF"
```

### 查看结构

```bash
python pptx_editor.py report.pptx --inspect
python pptx_editor_com.py report.pptx --inspect
```

### 交互模式

```bash
python pptx_editor.py report.pptx --interactive
python pptx_editor_com.py report.pptx --interactive
```

### 导出图片（COM 版）

```bash
python pptx_editor_com.py report.pptx --export-images
```

## 支持的指令

### 文本操作
- `把标题改成"新标题"`
- `第3页副标题改成"hello"`
- `删除第2页的表格`

### 样式操作
- `第1页标题字号改成36`
- `标题大一点` / `正文小一点`
- `第2页标题加粗`
- `标题颜色改成红色`
- `字体改成微软雅黑`

### 定位方式
- 按类型：标题 / 副标题 / 正文 / 表格 / 图片
- 按位置：左上 / 右上 / 左下 / 右下 / 居中
- 按内容：`"包含这段文字的"` 元素

### COM 独有
- `给第1页标题添加动画淡入`
- `第2页切换效果淡化`
- `导出PDF`

## 测试

```bash
# 生成测试 PPT
python gen_test.py

# 运行测试
python pptx_editor.py test_report.pptx --inspect
python pptx_editor.py test_report.pptx "把标题改成「测试成功」"
```

## 架构

```
指令文本 → parse_intent() → [{action, slide, target, params}]
                                    ↓
                              find_shape() 定位元素
                                    ↓
                              modify / delete / animate
                                    ↓
                                save()
```

## Windows 本地测试

见 `demo_test.bat` — 一键运行所有测试用例。
