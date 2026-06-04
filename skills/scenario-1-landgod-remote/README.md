# 方案一：本地 LLM + 远程 Windows（通过 LandGod）

## 方案概述

LLM 运行在本地 Linux 服务器上，通过 LandGod 远程调用部署在 Windows 机器上的 `pptx_editor_com.py`，实现跨平台的 PowerPoint 编辑。适用于：开发环境在 Linux、Office 在远程 Windows 的团队。

---

## 架构图

```
┌─────────────────────────────────────────────────────────────────┐
│                        本地 Linux 服务器                         │
│                                                                 │
│  用户输入 ──→ LLM (本地/API) ──→ 解析意图 ──→ 生成 Python 代码  │
│                                       │                         │
│                                       ▼                         │
│                              LandGod Gateway                    │
│                              POST /api/tool_call                │
└──────────────────────────────┬──────────────────────────────────┘
                               │ HTTP/WebSocket
                               ▼
┌──────────────────────────────────────────────────────────────────┐
│                      远程 Windows 机器                           │
│                                                                  │
│  LandGod Worker ──→ 执行 Python 代码 ──→ pptx_editor_com.py     │
│                                              │                   │
│                                              ▼                   │
│                                     COM 接口 → PowerPoint.exe    │
│                                              │                   │
│                                              ▼                   │
│                                       编辑后的 .pptx             │
└──────────────────────────────────────────────────────────────────┘
```

---

## 前置条件

1. **远程 Windows 机器**
   - Windows 10/11 或 Windows Server
   - 已安装 Microsoft Office（含 PowerPoint）
   - 已安装 Python 3.8+ 及 `pywin32`
   - 已部署 `pptx_editor_com.py`
   - LandGod Worker 已安装并连接到 Gateway

2. **本地 Linux 服务器**
   - LandGod Gateway 运行中
   - LLM 可用（本地 Ollama 或云端 API）
   - `pptx_editor_llm.py` 用于意图解析

3. **网络**
   - Linux 与 Windows 之间网络可达
   - LandGod Gateway 端口开放（默认 8080）

---

## 部署步骤

### 1. Windows 端部署

```bash
# 安装依赖
pip install pywin32

# 部署编辑器
# 将 pptx_editor_com.py 复制到 Windows 机器
# 例如放在 C:\pptx-editor\pptx_editor_com.py

# 安装并启动 LandGod Worker
landgod-worker --gateway http://linux-server:8080 --name win-office
```

### 2. Linux 端配置

```bash
# 确保 LandGod Gateway 已运行
landgod-gateway --port 8080

# 验证 Worker 已连接
curl http://localhost:8080/api/workers
# 应看到 "win-office" 在线
```

### 3. 上传 PPTX 文件到 Windows

```bash
# 通过 LandGod 文件传输
curl -X POST http://localhost:8080/api/file/upload \
  -F "worker=win-office" \
  -F "path=C:\\pptx-editor\\input.pptx" \
  -F "file=@./input.pptx"
```

### 4. 执行编辑命令

```bash
# 通过 LandGod tool_call 发送 Python 代码
curl -X POST http://localhost:8080/api/tool_call \
  -H "Content-Type: application/json" \
  -d '{
    "worker": "win-office",
    "tool": "python",
    "args": {
      "code": "import sys; sys.path.insert(0, r\"C:\\pptx-editor\"); from pptx_editor_com import PPTXEditorCOM; editor = PPTXEditorCOM(r\"C:\\pptx-editor\\input.pptx\"); editor.set_text(1, 1, \"新标题\"); editor.set_font_color(1, 1, 255, 0, 0); editor.save(); editor.close()"
    }
  }'
```

### 5. 下载编辑后的文件

```bash
curl -X POST http://localhost:8080/api/file/download \
  -H "Content-Type: application/json" \
  -d '{"worker":"win-office","path":"C:\\pptx-editor\\input.pptx"}' \
  -o output.pptx
```

---

## 使用示例

### 结合 LLM 的完整流程

```python
import requests
import json

GATEWAY = "http://localhost:8080"
WORKER = "win-office"
PPTX_PATH = r"C:\pptx-editor\report.pptx"

def llm_parse_intent(user_input):
    """使用 LLM 将自然语言转为编辑指令"""
    # 调用 pptx_editor_llm.py 的系统提示词解析意图
    # 返回 JSON actions 列表
    response = call_llm(
        system_prompt=PPTX_EDITOR_SYSTEM_PROMPT,
        user_message=user_input
    )
    return json.loads(response)

def execute_on_remote(code):
    """通过 LandGod 在远程 Windows 上执行代码"""
    resp = requests.post(f"{GATEWAY}/api/tool_call", json={
        "worker": WORKER,
        "tool": "python",
        "args": {"code": code}
    })
    return resp.json()

# 用户输入
user_input = "把第一页标题改成'年度报告'，红色加粗，36号字"

# LLM 解析
actions = llm_parse_intent(user_input)
# actions = [
#   {"method": "set_text", "args": [1, 1, "年度报告"]},
#   {"method": "set_font", "args": [1, 1, null, 36, true, false, [255,0,0]]}
# ]

# 生成代码
code = f"""
from pptx_editor_com import PPTXEditorCOM
editor = PPTXEditorCOM(r"{PPTX_PATH}")
editor.set_text(1, 1, "年度报告")
editor.set_font(1, 1, None, 36, True, False, (255,0,0))
editor.save()
editor.close()
"""

# 远程执行
result = execute_on_remote(code)
print(result)
```

### 使用 bat 文件包装（推荐）

由于 LandGod 的 `python` tool 可能有路径问题，建议用 bat 包装：

```bat
:: C:\pptx-editor\run_edit.bat
@echo off
cd /d C:\pptx-editor
call python -c "%~1"
```

然后通过 LandGod 调用：

```json
{
  "worker": "win-office",
  "tool": "shell",
  "args": {
    "command": "C:\\pptx-editor\\run_edit.bat \"from pptx_editor_com import PPTXEditorCOM; e=PPTXEditorCOM(r'C:\\pptx-editor\\input.pptx'); e.set_text(1,1,'Hello'); e.save(); e.close()\""
  }
}
```

---

## 注意事项 / 已知问题

### ⚠️ LandGod 4000 字符限制

LandGod 的 tool_call 参数有 **4000 字符限制**。对于复杂编辑：
- 将代码写入 `.py` 文件再执行
- 或分多次 tool_call

```python
# 先上传脚本
upload_file("edit_script.py", code_content)
# 再执行
execute_on_remote("exec(open(r'C:\\pptx-editor\\edit_script.py').read())")
```

### ⚠️ Session 0 问题

如果 LandGod Worker 作为 Windows 服务运行，会在 Session 0 中，COM 操作会失败。

**解决方案**：确保 Worker 在用户桌面会话中运行（手动启动或自动登录后启动）。

### ⚠️ BGR 颜色

LLM 生成的颜色值是 RGB 格式，`pptx_editor_com.py` 内部会自动转换。但如果直接写 COM 代码，需手动转换：

```python
# RGB(255,0,0) 红色 → BGR = 0x0000FF
font.Color.RGB = 0x0000FF  # 红色
```

### ⚠️ bat 文件中的 python 调用

Windows bat 文件中必须用 `call python` 而非直接 `python`，否则 bat 执行会提前终止：

```bat
:: ❌ 错误
python script.py

:: ✅ 正确
call python script.py
```

---

## 文件清单

| 文件 | 位置 | 说明 |
|------|------|------|
| `pptx_editor_com.py` | Windows 机器 | COM 编辑器核心，70+ 方法 |
| `pptx_editor_llm.py` | Linux 服务器 | LLM 意图解析，生成编辑指令 |
| `demo_test.py` | Windows 机器（可选） | 测试套件，验证 COM 环境 |
