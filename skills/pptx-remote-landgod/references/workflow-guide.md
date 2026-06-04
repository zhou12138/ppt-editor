# LandGod 远程 PPTX 编辑完整工作流指南

## 概述

本指南详细说明如何通过 LandGod 网关远程编辑 PowerPoint 文件的完整流程。

## 前置准备

### 1. LandGod Gateway 配置

确保 LandGod Gateway 已启动并监听 `localhost:8081`：

```bash
# 验证 Gateway 可用
curl http://localhost:8081/health
```

### 2. Windows Worker 环境

在 Windows Worker 上确认以下组件：

```powershell
# 检查 Python
python --version

# 检查 pywin32
python -c "import win32com.client; print('pywin32 OK')"

# 检查 Office
python -c "import win32com.client; app=win32com.client.Dispatch('PowerPoint.Application'); print(app.Version); app.Quit()"

# 确认 pptx_editor_com.py 可导入
python -c "from pptx_editor_com import PowerPointCOM; print('OK')"
```

### 3. 桌面会话要求

**重要：** Windows Worker 必须通过 RDP 保持桌面会话。Session 0（服务模式）无法执行 COM 的 `Open` 和 `SaveAs` 操作。

建议使用 `tscon` 或保持 RDP 连接不断开。

## 完整工作流示例

### 场景：修改一份销售报告的标题和图表颜色

#### 步骤 1：获取 PPTX 结构信息

```python
# 构造 inspect 脚本
script = """
from pptx_editor_com import PowerPointCOM
p = PowerPointCOM()
p.open("C:\\\\Reports\\\\Q4_Sales.pptx")
result = p.inspect()
print(result)
p.cleanup()
"""
```

通过 LandGod 发送：

```bash
ENCODED=$(echo -n "$script" | base64)
curl -X POST http://localhost:8081/api/tool_call \
  -H "Authorization: Bearer $TOKEN" \
  -d "{\"command\": \"python\", \"args\": \"$ENCODED\"}"
```

返回结果类似：

```json
{
  "slides": [
    {
      "index": 1,
      "shapes": [
        {"index": 1, "type": "TextBox", "text": "Q4 Sales Report", "left": 100, "top": 50},
        {"index": 2, "type": "Chart", "left": 100, "top": 200}
      ]
    }
  ]
}
```

#### 步骤 2：LLM 解析用户意图

用户说："把标题改成'2024年第四季度销售报告'，用红色加粗"

`pptx_editor_llm.py` 解析为：

```json
[
  {"action": "set_text", "slide": 1, "shape_index": 1, "text": "2024年第四季度销售报告"},
  {"action": "set_font_bold", "slide": 1, "shape_index": 1, "bold": true},
  {"action": "set_font_color", "slide": 1, "shape_index": 1, "color": "0000FF"}
]
```

> **注意：** 颜色使用 BGR 格式！红色 RGB `FF0000` 在 COM 中为 BGR `0000FF`。

#### 步骤 3：生成并发送执行脚本

```python
script = """
from pptx_editor_com import PowerPointCOM
p = PowerPointCOM()
p.open("C:\\\\Reports\\\\Q4_Sales.pptx")
p.set_text(1, 1, "2024年第四季度销售报告")
p.set_font_bold(1, 1, True)
p.set_font_color(1, 1, "0000FF")
p.save()
p.cleanup()
print("编辑完成")
"""
```

#### 步骤 4：验证结果

再次调用 `inspect()` 确认修改已生效。

## 常见问题排查

### 4000 字符限制

当操作较多时，将脚本拆分为多次调用：

```python
# 第一次调用：修改文本
script_1 = "p.open(...); p.set_text(1,1,'...'); p.save()"

# 第二次调用：修改格式
script_2 = "p.open(...); p.set_font_color(1,1,'0000FF'); p.save()"
```

### 内存不足（2GB 机器）

```python
# 在脚本开头添加内存清理
import subprocess
subprocess.run(["taskkill", "/f", "/im", "POWERPNT.EXE"], capture_output=True)
import gc; gc.collect()
```

### find_shape() 精确定位

当 shape_index 不确定时，使用 `find_shape()` 按类型/位置/文本匹配：

```python
shape = p.find_shape(slide_index=1, text_match="Sales Report")
```

## 安全注意事项

- LandGod Token 不要硬编码在脚本中，使用环境变量
- 敏感文件路径避免在日志中打印
- 编辑完成后及时调用 `cleanup()` 释放 COM 对象
