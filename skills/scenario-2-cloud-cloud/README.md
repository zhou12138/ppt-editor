# 方案二：云端 LLM + 云端 Windows（同机部署）

## 方案概述

LLM 使用云端 API（如 OpenAI、通义千问），PPTX 编辑在同一台 Windows 云服务器上完成。`pptx_editor_llm.py` 作为中枢，接收自然语言指令，调用 LLM API 解析意图，然后通过 `pptx_editor_com.py` 执行 COM 操作。

**最简单的部署方案**——一台 Windows 机器搞定一切。

---

## 架构图

```
┌─────────────────────────────────────────────────────────────┐
│                    Windows 云服务器                           │
│                                                              │
│  用户输入（命令行/API）                                       │
│       │                                                      │
│       ▼                                                      │
│  pptx_editor_llm.py                                         │
│       │                                                      │
│       ├──→ 云端 LLM API ──→ 返回 JSON 编辑指令               │
│       │    (OpenAI / 通义 / DeepSeek)                        │
│       │                                                      │
│       ▼                                                      │
│  pptx_editor_com.py                                         │
│       │                                                      │
│       ▼                                                      │
│  COM 接口 ──→ PowerPoint.exe ──→ 编辑 .pptx                 │
│                                                              │
└─────────────────────────────────────────────────────────────┘
         ▲                              │
         │    互联网 (HTTPS)            │
         ▼                              ▼
   ┌──────────┐                  ┌──────────┐
   │ LLM API  │                  │ 编辑后的  │
   │ 服务器    │                  │  .pptx   │
   └──────────┘                  └──────────┘
```

---

## 前置条件

1. **Windows 云服务器**（如 Azure VM、AWS EC2）
   - Windows 10/11 或 Windows Server 2019+
   - 已安装 Microsoft Office（含 PowerPoint）
   - Python 3.8+
   - `pip install pywin32 openai requests`

2. **LLM API 访问**
   - OpenAI API Key 或兼容的 API（通义千问、DeepSeek 等）
   - 网络可访问 API 端点

3. **桌面会话**（关键！）
   - 必须通过 RDP 登录，保持桌面会话
   - 或配置自动登录 + tscon 保活

---

## 部署步骤

### 1. 安装依赖

```powershell
pip install pywin32 openai requests
```

### 2. 部署项目文件

```powershell
# 将以下文件放到工作目录，如 C:\pptx-editor\
# - pptx_editor_com.py
# - pptx_editor_llm.py
```

### 3. 配置环境变量

```powershell
# OpenAI
set OPENAI_API_KEY=sk-xxxxxxxxxxxxx
set OPENAI_API_BASE=https://api.openai.com/v1

# 或通义千问
set OPENAI_API_KEY=sk-xxxxxxxxxxxxx
set OPENAI_API_BASE=https://dashscope.aliyuncs.com/compatible-mode/v1

# 或 DeepSeek
set OPENAI_API_KEY=sk-xxxxxxxxxxxxx
set OPENAI_API_BASE=https://api.deepseek.com/v1
```

### 4. 验证 COM 环境

```powershell
cd C:\pptx-editor
python demo_test.py
# 应生成 demo_output.pptx，无报错
```

### 5. 开始使用

```powershell
python pptx_editor_llm.py report.pptx "把标题改成红色大字"
```

---

## 使用示例

### 单次编辑

```bash
# 修改标题
python pptx_editor_llm.py report.pptx "把第一页标题改成'2024年度总结'，字号40，加粗，红色"

# 添加内容
python pptx_editor_llm.py report.pptx "在第二页添加一个表格，3行4列，内容是Q1到Q4的销售数据"

# 导出
python pptx_editor_llm.py report.pptx "导出为PDF"
```

### 交互模式

```bash
python pptx_editor_llm.py report.pptx --interactive

> 把所有标题字体改成微软雅黑
✅ 已修改 5 个标题的字体

> 第三页背景改成浅蓝色
✅ 已设置第3页背景色 RGB(173,216,230)

> 给第一页标题加个飞入动画
✅ 已添加飞入动画效果

> exit
已保存并退出。
```

### 作为 API 服务

```python
# 简易 Flask API 包装
from flask import Flask, request, jsonify, send_file
from pptx_editor_llm import PPTXEditorLLM

app = Flask(__name__)

@app.route("/api/edit", methods=["POST"])
def edit_pptx():
    data = request.json
    pptx_path = data["file"]
    instruction = data["instruction"]
    
    editor = PPTXEditorLLM(pptx_path)
    result = editor.execute(instruction)
    editor.save()
    editor.close()
    
    return jsonify({"status": "ok", "result": result})

@app.route("/api/download/<filename>")
def download(filename):
    return send_file(f"C:/pptx-editor/{filename}")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
```

```bash
# 调用 API
curl -X POST http://your-server:5000/api/edit \
  -H "Content-Type: application/json" \
  -d '{"file":"C:\\pptx-editor\\report.pptx","instruction":"把标题改成蓝色"}'
```

---

## 注意事项 / 已知问题

### ⚠️ 桌面会话（最重要）

COM 操作 **必须** 在有桌面的会话中运行。如果通过 SSH 或断开 RDP 后运行，PowerPoint COM 会失败。

**保活方案**：

```bat
:: 断开 RDP 前执行，保持桌面会话
tscon %sessionname% /dest:console
```

或设置 Windows 自动登录：

```reg
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon]
"AutoAdminLogon"="1"
"DefaultUserName"="your_user"
"DefaultPassword"="your_password"
```

### ⚠️ API 费用

每次编辑都会调用 LLM API，注意费用控制：
- GPT-4 约 $0.03-0.06/次编辑
- DeepSeek 约 ¥0.01/次编辑
- 通义千问 qwen-plus 约 ¥0.004/次编辑

### ⚠️ 延迟

云端 API 调用延迟约 1-5 秒，加上 COM 操作时间，单次编辑总耗时约 3-10 秒。

### ⚠️ 并发

PowerPoint COM 不支持多线程并发操作。如需处理多个文件，使用队列串行执行。

---

## 文件清单

| 文件 | 说明 |
|------|------|
| `pptx_editor_com.py` | COM 编辑器核心，70+ 方法 |
| `pptx_editor_llm.py` | LLM 集成层，自然语言→编辑指令 |
| `demo_test.py` | 测试套件，验证环境 |
