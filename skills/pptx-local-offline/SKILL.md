---
name: pptx-local-offline
description: "Edit PowerPoint completely offline on a local Windows machine. Three modes: (A) local LLM via Ollama/LM Studio, (B) Claude/AI agent direct script execution, (C) Claude/AI agent JSON actions dispatch. Zero cloud dependency for mode A; mode B/C work with any local AI agent. Triggers on: offline pptx, local llm powerpoint, air-gapped pptx editing, claude pptx, local agent pptx."
---

# 本地离线 PowerPoint 编辑

## 概述

在同一台 Windows 机器上执行 PowerPoint COM 自动化，实现离线 PPTX 编辑。支持三种执行模式，按需选择。

## 使用优先级

默认按以下优先级选择模式：

1. `JSON actions`
2. `exec-script`
3. 其他模式（LLM 自然语言 / 交互模式）

选择原则：
- 能用 JSON actions 表达的需求，优先用 `--exec-actions`
- JSON actions 还不支持、或需要复杂控制流时，再退到 `--exec-script`
- 只有在前两者都不合适时，才使用自然语言解析或交互模式

发生真实修改后，PowerPoint 会额外保留 30 秒再关闭，便于确认结果。

## 三种执行模式

### 模式 A：本地 LLM 意图解析（Ollama/LM Studio）

LLM 运行在本地，解析自然语言指令为 JSON actions 再执行 COM。完全离线，零云端依赖。

```
用户指令 → 本地 LLM (Ollama) → JSON Actions → pptx_editor_com.py → COM → PowerPoint
```

```bash
set OPENAI_API_BASE=http://localhost:11434/v1
set OPENAI_API_KEY=ollama
python pptx_editor_llm.py deck.pptx "把标题改成红色加粗"
python pptx_editor_llm.py deck.pptx --interactive
```

### 模式 B：本地 AI Agent 脚本执行 `--exec-script`（推荐 Claude Code）

本地 Claude Code 等 AI Agent 直接生成 Python 脚本，脚本内可用 `ppt`（已打开的 PowerPointCOM 实例）和 `filepath` 变量。**无需任何 API Key，无需 Ollama。**

底层执行策略现在支持通过 `--backend pywin32|vba` 切换。默认是 `pywin32`。`vba` 策略会通过 `Application.Run` 调用 VBA 桥接宏，适合已有宏工程或准备把动作下沉到 VBA 的场景。当前 `vba` 策略推荐搭配 `--exec-actions` / `--interactive-actions` 使用；`--exec-script` 仍然只支持 Python。

`inspect()` 现在会额外标记非文本内容：占位符或普通 shape 里如果包含图片、图表、表格或媒体，会在结构里给出 `has_image`、`has_chart`、`has_table`、`has_media`，CLI 输出也会显示 `[图片]`、`[图表]` 之类的摘要，不再把这类元素误显示成“(无)”。

现在脚本模式额外内建了两个辅助函数：
- `log_note("文本", slide=None, append=True)`：把执行进度写到演讲者备注
- `sleep(seconds, slide=None, note=None, append=True)`：可选先写备注，再等待若干秒

```bash
python pptx_editor_llm.py deck.pptx --exec-script edit.py
python pptx_editor_llm.py deck.pptx --exec-script edit.py --headed
python pptx_editor_llm.py deck.pptx --exec-script edit.py --notes-progress --note-slide 1
python pptx_editor_llm.py deck.pptx --interactive-script --headed
python pptx_editor_llm.py deck.pptx --exec-script edit.py --output out.pptx
python pptx_editor_llm.py deck.pptx --inspect --exec-script edit.py
```

Claude 生成的脚本示例：
```python
# edit.py — ppt 和 filepath 已注入为全局变量
log_note("开始处理第一页标题", slide=1, append=False)
structure = ppt.inspect()
shapes = ppt.find_shape(1, {"type": "title"})
for s in shapes:
    ppt.modify_font(s, bold=True, color=0x0000FF)  # BGR红色
sleep(1.5, slide=1, note="标题已加粗并改为红色")
ppt.add_animation(1, shapes[0], "fade")
```

脚本会话模式会保持同一个 COM 会话，持续从 stdin 读取命令。支持：
- `help`
- `inspect`
- `status`
- `save`
- `saveas out.pptx`
- `quit`
- `close`
- 直接输入脚本路径
- `{"command":"script","path":"edit.py"}`
- `{"command":"script_inline","code":"ppt.set_notes(1, 'hello')"}`

### 模式 C：本地 AI Agent JSON 动作执行 `--exec-actions`

Claude 等 AI Agent 生成 JSON actions 数组，跳过 LLM 解析直接 dispatch 执行。接受 JSON 字符串或 .json 文件路径。**无需任何 API Key，无需 Ollama。**

```bash
python pptx_editor_llm.py deck.pptx --exec-actions '[{"action":"modify_font","slide":1,"target":{"type":"title"},"params":{"bold":true}}]'
python pptx_editor_llm.py deck.pptx --exec-actions actions.json
python pptx_editor_llm.py deck.pptx --interactive-actions --headed
python pptx_editor_llm.py deck.pptx --exec-actions actions.json --headed --notes-progress
python pptx_editor_llm.py deck.pptx --exec-actions actions.json --dry-run
```

JSON actions 也支持等待动作：

```json
[
    {"action": "modify_font", "slide": 1, "target": {"type": "title"}, "params": {"bold": true}},
    {"action": "sleep", "params": {"seconds": 2}},
    {"action": "animation", "slide": 1, "target": {"type": "title"}, "params": {"effect": "fade"}}
]
```

JSON 会话模式会保持同一个 COM 会话，持续从 stdin 读取命令。推荐使用 JSON command envelope：

```json
{"command":"actions","payload":[{"action":"set_slide_background_image","slide":1,"params":{"image_path":"bg.png"}}]}
{"command":"inspect"}
{"command":"save"}
{"command":"quit"}
```

也支持直接输入 JSON actions 数组或单个 action 对象，以及以下简命令：
- `help`
- `inspect`
- `status`
- `save`
- `saveas out.pptx`
- `quit`
- `close`

完整示例见 [interactive-session-guide.md](./references/interactive-session-guide.md)。

### 模式选择指南

| 场景 | 推荐模式 | 原因 |
|------|----------|------|
| 气隙/涉密环境，无任何外部 AI | A (Ollama) | 完全自包含 |
| 本地跑 Claude Code/Cursor | C (exec-actions) | 优先使用结构化 JSON，行为稳定且易审计 |
| JSON 还不能覆盖的复杂操作 | B (exec-script) | 作为第二优先级，适合复杂控制流和尚未接入的 COM 能力 |
| Agent 通过 MCP/API 远程调用 | C (exec-actions) | JSON 标准化接口，易于集成 |
| 简单批量修改 | C (exec-actions) | JSON 声明式，易于复用和版本化 |

### 通用选项

| 选项 | 说明 | 适用模式 |
|------|------|----------|
| `--inspect` | 打印 PPTX 结构 | 全部 |
| `--output / -o` | 指定输出文件路径 | 全部 |
| `--dry-run` | 仅解析不执行 | A, C |
| `--api-base` | 覆盖 API 端点 | A |
| `--model` | 覆盖模型名 | A |
| `--api-key` | 覆盖 API 密钥 | A |
| `--interactive-actions [JSON]` | 在同一 COM 会话中持续执行 JSON 命令 | C |
| `--interactive-script [SCRIPT]` | 在同一 COM 会话中持续执行脚本命令 | B |
| `--headed` | 以可见窗口模式打开 PowerPoint | 全部 |
| `--notes-progress` | 自动把当前指令或脚本进度写到演讲者备注 | A, B, C |
| `--note-slide` | 将进度备注固定写到指定页 | A, B, C |
| `--backend` | 选择底层策略：`pywin32` 或 `vba` | B, C |
| `--vba-module` | 指定 VBA 桥接宏模块名，默认 `PptEditorBridge` | B, C |

## 安装配置（所有模式通用）

```powershell
pip install pywin32
# 仅模式 A 需要：
pip install requests
```

### 模式 A 额外配置：安装 Ollama

```powershell
winget install Ollama.Ollama
ollama pull qwen2.5:7b

setx OPENAI_API_BASE "http://localhost:11434/v1"
setx OPENAI_API_KEY "ollama"
```

## 模型推荐（仅模式 A）

| 模型 | 内存需求 | JSON 输出质量 | 中文支持 | 推荐场景 |
|------|----------|--------------|----------|----------|
| `qwen2.5:7b` | 8GB | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | **推荐默认** |
| `qwen2.5:14b` | 16GB | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | 高质量需求 |
| `qwen2.5:3b` | 4GB | ⭐⭐⭐ | ⭐⭐⭐⭐ | 低配机器 |
| `llama3.1:8b` | 8GB | ⭐⭐⭐⭐ | ⭐⭐⭐ | 英文场景 |

> **关键：** 模式 A 需要 7B+ 模型才能可靠输出 JSON。模式 B/C 不受此限制。

## 脚本文件

| 文件 | 用途 |
|------|------|
| `pptx_editor_com.py` | COM 引擎，所有 PowerPoint 操作方法 |
| `pptx_editor_llm.py` | 统一入口（三种模式） |
| `pptx_editor.py` | 纯规则引擎，无需 LLM 的离线备选方案 |

## 注意事项

| 问题 | 说明 |
|------|------|
| **BGR 颜色** | COM 使用 BGR 格式！红=0x0000FF，蓝=0xFF0000 |
| **1-Based 索引** | 所有 COM 索引从 1 开始 |
| **Session 0 限制** | schtasks 启动的进程无法 Open/SaveAs，需 RDP 桌面会话 |
| **自动延迟关闭** | 发生真实修改后，PowerPoint 会保留 30 秒再关闭，便于人工确认结果 |
| **Notes 进度显示** | `--notes-progress` 默认会覆盖目标备注页当前内容；如需保留历史请在脚本里用 `log_note(..., append=True)` |
| **内存竞争（模式 A）** | Ollama + PowerPoint 同时运行，建议 16GB+ |
| **模型质量（模式 A）** | 本地模型不如 GPT-4/Claude，复杂指令可能需多次尝试 |
| **VBA 策略要求** | 需先把 `references/PptEditorBridge.bas` 导入到启用宏的演示文稿或 add-in，并补齐桥接函数 |

## 适用场景

- 🔒 涉密/气隙（air-gapped）环境
- 🏢 企业内网无外网访问
- 🤖 本地 AI Agent（Claude Code、Cursor）直接编辑 PPTX
- 💰 无需 API 费用的长期使用

## 参考文档

- `references/setup-guide.md` - Ollama 安装与模型选择详细指南
- `references/interactive-session-guide.md` - 交互会话命令协议
- `references/PptEditorBridge.bas` - VBA backend 桥接模块模板
