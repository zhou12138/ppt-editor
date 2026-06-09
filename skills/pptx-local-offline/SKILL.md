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

底层执行策略现在支持通过 `--backend pywin32|vba|csharp|csharp-codeact|csharp-addin|pywin32-addin` 切换。默认是 `pywin32`。六种策略主要用于性能对比，证明瓶颈在跨进程 IPC 而非语言本身：
- `pywin32`：跨进程后期绑定 COM（基线，Python × 跨进程）。
- `vba`：进程内 VBA 桥接宏（`Application.Run`），最快，进程内参照基准。
- `csharp`：跨进程常驻 C# host（exe，JSON-RPC），与 pywin32 同机制但换 C#（C# × 跨进程）。
- `csharp-codeact`：与 `csharp` 共用同一个常驻 host，但把每个动作翻译成 C# 脚本经 Roslyn(`execute_code`)执行；可把整批动作编译成**一段脚本一次往返**（CodeAct 模式，见下文）。
- `csharp-addin`：**进程内** C# COM 加载项，运行在 POWERPNT.EXE 内（与 VBA 同机制，早绑定 interop），需先注册（见下文）（C# × 进程内）。
- `pywin32-addin`：**进程内** Python（pywin32）COM 加载项，同样运行在 POWERPNT.EXE 内，复用同一套 `PowerPointCOM` 引擎，需先注册（见下文）（Python × 进程内）。

> 矩阵说明：`pywin32`（Python/跨进程）与 `pywin32-addin`（Python/进程内）是同语言、仅执行位置不同的对照组；`csharp` 与 `csharp-addin` 同理。两组同时变快即可隔离出真正的变量——是跨进程 COM marshalling（IPC）而非 Python/C# 语言差异拖慢了 `pywin32`。

`vba` 策略会通过 `Application.Run` 调用 VBA 桥接宏，适合已有宏工程或准备把动作下沉到 VBA 的场景。当前 `vba` / `csharp` / `csharp-codeact` / `csharp-addin` / `pywin32-addin` 策略推荐搭配 `--exec-actions` / `--interactive-actions` 使用；`--exec-script` 仍然只支持 Python。

> 两个进程内加载项（`csharp-addin` / `pywin32-addin`）与 VBA 一样，只在**交互式桌面会话**启动的 PowerPoint 里加载（PowerPoint 不会在无人值守 / 服务会话里加载加载项）。注册见下文「进程内加载项注册」。


`vba` 策略现在依赖 skill 内置的 `references/JsonConverter.bas`。它不是可选工具，而是 VBA backend 的 JSON 桥接组件：
- `ExecuteActionJson()` 用它把 Python 传来的 action JSON 解析成 VBA 可操作的 `Dictionary` / `Collection`
- `InspectPresentationJson()` 用它把 VBA 侧组装好的结构序列化成 JSON 再返回给 Python

不要再从外部下载通用版 `VBA-JSON` 覆盖它。skill 自带的是一份为 PowerPoint VBA 场景裁剪过的精简版本，用来避开社区通用版在 PowerPoint 环境下的兼容性和卡死问题。

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
| `--backend` | 选择底层策略：`pywin32` / `vba` / `csharp` / `csharp-codeact` / `csharp-addin` / `pywin32-addin` | B, C |
| `--vba-module` | 指定 VBA 桥接宏模块名，默认 `PptEditorBridge` | B, C |

## Backend 性能对比

基于 7 页真实 PPT 的本地实测数据（11 项操作，Windows 11 + Office 16 + Python 3.12）：

> 注：inspect 性能与 **shape 总数**强相关，而非页数。上表基于 7 页 / 约几十个 shape；在 119 shapes 的大文件上 pywin32 inspect 可达 20s+，VBA 仍保持亚秒级。评估时请以 shape 数为准。

| 操作 | pywin32 | VBA | VBA 加速比 |
|------|---------|-----|-----------|
| inspect (7页结构) | 3.90s | 0.20s | **20x** |
| inspect ×5 | 20.13s | 1.00s | **20x** |
| find 7 个标题 | 0.35s | 0.21s | 1.7x |
| modify_font (单个) | 0.11s | 0.04s | 2.7x |
| modify_font (7页) | 0.90s | 0.12s | **7.6x** |
| modify_text | 0.05s | 0.02s | 2.6x |
| move_shape | 0.08s | 0.03s | 3.1x |
| resize_shape | 0.02s | 0.01s | 1.6x |
| add + delete textbox | 0.20s | 0.04s | **5.2x** |
| set_notes | 0.10s | 0.02s | **6.4x** |
| batch 17 混合操作 | 1.28s | 0.37s | **3.5x** |
| **总计 (不含 open)** | **27.1s** | **2.0s** | **⚡ 13.3x** |

资源占用：

| 指标 | pywin32 | VBA |
|------|---------|-----|
| CPU user | 4.09s | 0.09s (45x 更低) |
| CPU kernel | 9.66s | 0.08s (120x 更低) |
| Python 内存 | 28.9 MB | 28.3 MB (持平) |
| PowerPoint 内存 | 373.6 MB | 378.0 MB (持平) |

### 性能差异原因

- **pywin32**：每次属性访问都是跨进程 COM IPC 调用（进程间通信），inspect 遍历所有 shape 属性时开销巨大
- **VBA**：在 PowerPoint 进程内直接访问对象模型，零 IPC 开销；JSON 序列化后一次性返回给 Python

### 选择建议

| 场景 | 推荐 backend |
|------|-------------|
| 频繁 inspect / 大量读操作 | `vba` (20-33x 更快) |
| 批量修改多页 slide | `vba` (3-9x 更快) |
| 单次简单修改 | `pywin32` (差距小，无需导入 VBA 模块) |
| 无法启用 VBA 宏信任 | `pywin32` (不依赖 VBA) |
| 需要扩展自定义 COM 操作 | `pywin32` (Python 直接写，更灵活) |

### 性能 Insights

1. **IPC 是 pywin32 的致命瓶颈**：pywin32 每次访问 shape 属性都是跨进程 COM 调用（~20ms/次），inspect 7 页需 ~200 次往返。VBA 在 PowerPoint 进程内直接访问，零 IPC，inspect 快 27 倍。

2. **读密集收益最大，混合负载趋于持平**：纯读(inspect) 快 27-33x，纯写(modify) 快 3-9x，但混合高频循环 ~1x — 此时瓶颈已转移到 PowerPoint 进程内部的对象遍历和渲染。

3. **VBA 延时恒定 ~20ms/op**：不随 batch 大小增长。pywin32 在 116-173ms 间波动，COM proxy 缓存和进程调度开销随并发放大。

4. **CPU 差距比延时差距更大**：延时快 12x，但 Python CPU 低 50x。pywin32 大量 CPU 花在 COM marshalling（序列化、跨进程内存拷贝），VBA 将计算下沉到 PowerPoint 进程，Python 端近乎零开销。

5. **内存无差异**：VBA 模块仅增加 ~23MB（PowerPoint 侧），Python 侧完全一致。性能提升是"免费的"。

6. **真实交互体感**：典型用户流程 `inspect → 3~5 个 action → 保存`，pywin32 约 5.7s，VBA 约 0.3s — 从"等一等"到"瞬间完成"。

## C# Interop backend（`--backend csharp`）

第三种执行策略：一个常驻的 C# 进程用**后期绑定 COM（dynamic / IDispatch，与 pywin32 同机制）**驱动 PowerPoint，
通过 stdin/stdout 的 **行分隔 JSON-RPC** 与 Python 通信。主要用于 pywin32 / VBA / C# 四方性能对比（另一路是进程内的 `csharp-addin`，见下节）。

> 它是一个**可执行 host（exe），不是注册的 COM dll** —— 无需 regsvr32，不写注册表。

### 其他 agent / 调用方如何使用

**方式 1（推荐）：经 Python backend，自动拉起 exe**

```powershell
# 先构建一次（仓库内开发）：
dotnet build csharp_interop/PptInteropHost -c Release
# 之后正常用统一入口，加 --backend csharp 即可：
python scripts/pptx_editor_llm.py deck.pptx --inspect --backend csharp
```

Python 侧 `PowerPointCSharp` 会自动定位 exe，查找顺序：
1. 环境变量 `PPTX_EDITOR_CSHARP_HOST`（指向 `PptInteropHost.exe` 的绝对路径）
2. 从 `ppt_backend.py` 向上逐层查找 `csharp_host/PptInteropHost.exe`（**安装版 skill**：`install-local-offline-skill.ps1` 会把 host 发布到 `<skill>/csharp_host/`）
3. 向上查找 `csharp_interop/PptInteropHost/bin/<Release|Debug>/<tfm>/PptInteropHost.exe`（**仓库开发**构建产物）

> 安装版 skill（`~/.copilot/skills/pptx-local-offline`）上层没有 `csharp_interop/`，
> 因此安装脚本会在安装时执行 `dotnet publish` 把 host 放进 `csharp_host/`；
> 若目标机无 dotnet，可改用 `PPTX_EDITOR_CSHARP_HOST` 指向已有 exe（`-SkipCSharp` 可跳过发布）。

**方式 2：直接与 exe 通信（语言无关）**

任何语言都可以 spawn 这个 exe，按行写入 JSON 请求、按行读回 JSON 响应：

```text
→ {"cmd":"ping"}
← {"ok":true,"result":"pong"}
→ {"cmd":"open","path":"C:\\deck.pptx","visible":true}
← {"ok":true,"result":3}
→ {"cmd":"inspect"}
← {"ok":true,"result":{"slides":[...]}}
→ {"cmd":"execute_action","action":{"action":"modify_font","slide":1,"target":{"name":"Title"},"params":{"bold":true,"color":255}}}
← {"ok":true,"result":"加粗, 颜色→0xFF"}
→ {"cmd":"quit"}
← {"ok":true,"result":"bye"}
```

支持的 `cmd`：`ping` / `open` / `inspect` / `inspect_slide` / `execute_action` / `execute_code` / `set_notes` / `append_notes` / `save` / `quit`。
`execute_action` 的 `action` 负载与 VBA backend 完全一致（见模式 C 的 JSON 格式），target 定位语义对齐 pywin32（`type`/`name`/`text_match`/`position`/`index`/`id`）。

> ⚠️ 颜色仍是 **BGR**（红=255/0xFF）；索引 **1-based**；PowerPoint 不允许 `Application.Visible=false`，host 已自动用无窗口方式打开演示文稿。

### CodeAct 模式（`execute_code`）—— C# host 专属

C# host 在 `execute_action`（逐条 JSON）之外，还提供第二种**编排策略** `execute_code`：把整个编辑计划写成**一段 C# 脚本一次发送**，host 用 Roslyn（`Microsoft.CodeAnalysis.CSharp.Scripting`）在进程内编译执行，对着 `PptApi` 契约编程。这就是 CodeAct 模式——把 `model → tool → model` 循环塌缩成一次执行。

> `execute_action`（JSON）与 `execute_code`（CodeAct）是**同一个 host 上的两种编排策略**，区别只在调用方发哪种 `cmd`。CodeAct 本身不需要独立的 host：它复用 `--backend csharp` 那一个常驻进程、同一条 NDJSON 协议。为了方便 benchmark / CLI 对照，另提供了 `--backend csharp-codeact`（见下「调用方式 3」），它在**同一个 C# host** 上把每个 action 自动翻译成 `execute_code` 脚本。

| 维度 | `execute_action`（JSON） | `execute_code`（CodeAct） |
|------|--------------------------|---------------------------|
| 粒度 | 一条 = 一个原子操作 | 一段 = 整个计划 |
| 往返次数 | N 步 = N 次 stdin/stdout 往返 | N 步 = **1 次**往返 |
| 编排者 | 调用方/LLM（model→tool→model 循环） | 脚本本身（一次执行内完成） |
| 控制流 | 无（只能线性发命令） | 有：循环 / if / 过滤 / 聚合 / 中间变量 |
| 取值再用 | 难（inspect→模型解析→再发下一条） | 易（`var t=Title(1)` 直接拿对象继续用） |
| 契约 | JSON action 名 + params 字段 | `PptApi` 的 C# 方法签名（带类型/默认值） |
| 灵活度 | 受限于已实现的 action 列表 | 任意逻辑 + 经 `App`/`Prs` 直达原始 COM |

> 何时仍用 JSON：只有 1-2 步、每步需单独可见 / 审批、副作用要逐个确认时。

**调用方式 1：Python backend 的 `code_act()`**

`PowerPointCSharp` 暴露了 `code_act(script)` 便捷方法（仅 `csharp` backend 有），把脚本经 `execute_code` 发给 host，返回脚本里 `Print(...)` 的输出：

```python
from ppt_backend import create_backend
ppt = create_backend("csharp")          # 同一个 host 进程
ppt.open("deck.pptx")
out = ppt.code_act('''
    var t = Title(1);
    SetText(t, "新标题");
    SetFont(t, bold:true, colorBgr:0x0000FF);   // BGR 红
    AddTextbox(1, "副标题", 60, 160, 400, 50);
    Print($"slides={SlideCount}");
''')
ppt.save(); ppt.close()
```

> 注意：`code_act()` 是原生 Python / NDJSON 层的低阶 API，让你手写 C# 脚本。若只是想让标准 JSON action **走 CodeAct 路径**（而不手写 C#），用 `--backend csharp-codeact`（调用方式 3）更方便。

**调用方式 3：`--backend csharp-codeact`（把 JSON action 自动翻译成 CodeAct）**

`PowerPointCSharpCodeAct`（`PowerPointCSharp` 的子类）复用同一个 C# host，但把每个 JSON action 翻译成一段 `PptApi` C# 语句，经 `execute_code` 执行；`run_actions([...])` 会把一批可翻译的 action **拼成一段脚本、一次往返**（benchmark 里的 CodeAct 优势），遇到无法翻译的 action 自动降级回 JSON `execute_action`：

```powershell
# 与 --backend csharp 同一个 host；只是动作走 execute_code 而非逐条 JSON
python scripts/pptx_editor_llm.py deck.pptx --exec-actions actions.json --backend csharp-codeact
```

```python
from ppt_backend import create_backend
ppt = create_backend("csharp-codeact")
ppt.open("deck.pptx")
ppt.run_actions([                              # N 条 action → 1 次往返
    {"action": "modify_font", "slide": 1, "target": {"type": "title"}, "params": {"bold": True, "color": 255}},
    {"action": "add_textbox", "slide": 1, "params": {"text": "副标题", "left": 60, "top": 160, "width": 400, "height": 50}},
])
ppt.save(); ppt.close()
```

**调用方式 2：直接发 `execute_code`（语言无关）**

```text
→ {"cmd":"execute_code","code":"var t=Title(1); SetText(t,\"新标题\"); SetFont(t,bold:true,colorBgr:0x0000FF); Print(\"done\");"}
← {"ok":true,"result":{"output":"done\n"}}
```

**`PptApi` 契约（CodeAct 脚本可用 API）** —— 定义在 `PptApi.cs`。颜色 **BGR**、位置/大小单位 **points**（72pt=1in）、索引 **1-based**：

| 分类 | 成员 |
|------|------|
| 输出 | `Print(msg)`、`Output` |
| 导航 | `SlideCount`、`Slide(i)`、`Shape(slide,idx)`、`ShapeCount(slide)`、`FindByText(slide,contains)`、`FindByName(slide,name)`、`FindById(slide,id)`、`Title(slide)` |
| 文本/字体 | `SetText(shp,text)`、`GetText(shp)`、`SetFont(shp, size?, bold?, italic?, colorBgr?, name?)` |
| 几何 | `Move(shp,left,top)`、`Resize(shp,width,height)` |
| 外观 | `SetFill(shp,bgr)`、`SetBorder(shp,bgr,weight?)`、`SetSlideBackground(slide,bgr)` |
| 创建 | `AddTextbox(slide,text,left,top,width,height)` |
| 备注 | `SetNotes(slide,text)` |
| 原始 COM | `App`（Application）、`Prs`（Presentation） |

> CodeAct 减少的是**调用方↔host 的往返次数**，与「进程内 vs 进程外」是两个正交维度——它仍是进程外 COM host。更完整的协议/契约/驱动示例见 `csharp_interop/PptInteropHost/README.md`。

## C# in-process add-in backend（`--backend csharp-addin`）

第四种执行策略：一个 **进程内（in-process）的 C# COM 加载项**，被 PowerPoint 在启动时加载、运行在 `POWERPNT.EXE` 进程**内部**——和 VBA 同机制，**零跨进程 IPC**。
用于证明「VBA 的速度优势来自进程内 COM 访问，而非语言本身」：把语言固定为 C#、只改变进程边界（`csharp` exe host 是跨进程，`csharp-addin` 是进程内），即可隔离 IPC 这一变量。

> 它是一个**注册到 HKCU 的 COM dll 加载项**（不是 exe host）。PowerPoint **只在交互式启动时**加载加载项；自动化/`DispatchEx` 强制 connect 会被 PowerPoint 策略拒绝（E_ABORT）。因此 Python backend 会以交互方式拉起 `POWERPNT.EXE deck.pptx`，再通过 `Application.COMAddIns` 拿到桥接对象。

### 安装 / 注册（每用户，无需管理员）

`install-local-offline-skill.ps1` 会在安装时自动完成构建 + 部署 + 注册：

```powershell
# 默认：构建 net48 加载项，复制到 <skill>/csharp_addin/，并写入 HKCU 注册项（LoadBehavior=3）
powershell -File install-local-offline-skill.ps1

# 跳过所有 C# backend（exe host + 加载项）：
powershell -File install-local-offline-skill.ps1 -SkipCSharp

# 部署但不写注册表（之后手动注册）：
powershell -File install-local-offline-skill.ps1 -SkipAddinRegister
```

安装脚本把构建产物（`PptEditorAddin.dll` + `Newtonsoft.Json.dll`）放进 `<skill>/csharp_addin/`，并复制 `register.ps1` / `unregister.ps1` 到同目录。注册写入的 `CodeBase` 指向**部署后的 DLL 绝对路径**，因此在每台目标机上独立可用。

手动（重新）注册 / 反注册：

```powershell
powershell -File "<skill>\csharp_addin\register.ps1"   -DllPath "<skill>\csharp_addin\PptEditorAddin.dll"
powershell -File "<skill>\csharp_addin\unregister.ps1"
```

### 使用

```powershell
python scripts/pptx_editor_llm.py deck.pptx --inspect --backend csharp-addin
```

Python 侧 `PowerPointCSharpAddin` 通过 `PPTX_EDITOR_POWERPNT`（或自动探测 Office16 路径）定位 `POWERPNT.EXE`，交互式打开演示文稿后从 `Application.COMAddIns.Item("PptEditor.AddIn").Object` 取得进程内桥接对象。桥接方法（`InspectJson` / `InspectSlideJson` / `ExecuteActionJson`）的 `action` 负载与 VBA / C# host **完全一致**。

> ⚠️ 加载项依赖 **.NET Framework 4.8 运行时**（Office 机器通常自带）与 **dotnet SDK**（仅安装时构建用）。若目标机无 dotnet，可在有 dotnet 的机器构建后，把 `csharp_addin/` 整目录拷过去再跑 `register.ps1 -DllPath ...`。颜色仍是 **BGR**、索引 **1-based**。

> 早绑定说明：加载项已改为**早绑定 interop**（直接引用 GAC 中的 Office PIA 并 `EmbedInteropTypes=true`），不再用 `dynamic`。每个 PowerPoint 调用都在编译期类型检查，并走 vtable 调用，比 late-bound `dynamic` 更快、更安全；输出 DLL 自包含，无需随附 interop 程序集。

## Python in-process add-in backend（`--backend pywin32-addin`）

第五种执行策略：一个 **进程内（in-process）的 Python（pywin32）COM 加载项**，被 PowerPoint 在启动时加载、运行在 `POWERPNT.EXE` 进程**内部**——和 VBA / C# 加载项同机制，**零跨进程 IPC**。

这一格补全了 「Python / C#」×「跨进程 / 进程内」对照矩阵的最后一个单元：

| | 跨进程（out-of-proc） | 进程内（in-proc） |
|---|---|---|
| **Python** | `pywin32`（基线） | `pywin32-addin` |
| **C#** | `csharp`（exe host） | `csharp-addin` |
| **VBA** | — | `vba`（参照基准） |

把语言固定为 Python、只改变进程边界（`pywin32` 跨进程 vs `pywin32-addin` 进程内），即可与 C# 那一行平行地验证：**真正拖慢 `pywin32` 的是跨进程 COM marshalling（IPC），不是 Python 语言本身。**

> 它通过 pywin32 把自身注册为实现 `IDTExtensibility2` 的进程内 COM server，由 64 位 `POWERPNT.EXE` 在启动时加载（需 64 位 Python + pywin32）。它**复用同一套 `PowerPointCOM` 引擎**：加载项在 PowerPoint 进程内用 `ActivePresentation` 绑定当前演示文稿，因此不会再起第二个 PowerPoint。与 C# 加载项一样，只在**交互式桌面会话**加载。

### 安装 / 注册（每用户，无需管理员）

`install-local-offline-skill.ps1` 会在注册 C# 加载项之后，自动注册 Python 加载项（写入 HKCU `...\PowerPoint\Addins\PptEditor.PyAddIn`，LoadBehavior=3）。`-SkipAddinRegister` 会同时跳过两个加载项的注册。

手动（重新）注册 / 反注册：

```powershell
python "<skill>\scripts\pptx_pyaddin.py"              # 注册 COM server + HKCU 加载项
python "<skill>\scripts\pptx_pyaddin.py" --unregister # 反注册
python "<skill>\scripts\pptx_pyaddin.py" --clean      # 反注册并清理
```

### 使用

```powershell
python scripts/pptx_editor_llm.py deck.pptx --inspect --backend pywin32-addin
```

Python 侧 `PowerPointPywin32Addin`（继承自 `PowerPointCSharpAddin`）以交互方式拉起 `POWERPNT.EXE deck.pptx`，再从 `Application.COMAddIns.Item("PptEditor.PyAddIn").Object` 取得进程内桥接对象。桥接方法（`InspectJson()` / `InspectSlideJson(i)` / `ExecuteActionJson(json)`）的 `action` 负载与 VBA / C# 完全一致。

> ⚠️ 依赖 **64 位 Python + pywin32**。`set_notes` / `append_notes` 在加载项内直接走引擎方法；其余 action 走 `_dispatch`。颜色仍是 **BGR**、索引 **1-based**。

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
| **VBA 策略要求** | 需把 `references/PptEditorBridge.bas` 和 `references/JsonConverter.bas` 一起导入到启用宏的演示文稿或 add-in；`JsonConverter.bas` 是 VBA backend 的必需桥接组件 |

### 已知限制与踩坑

| 场景 | 说明 / 规避 |
|------|------|
| **OneDrive 同步路径** | `Presentations.Open()` 对 `C:\Users\xxx\OneDrive - .../xxx.pptx` 这类同步路径可能失败。规避：先把文件复制到本地非同步目录（如 `%TEMP%`）再打开。 |
| **无窗口模式被回收** | `visible=False` 下后台 PowerPoint 进程可能被系统/策略回收，后续 COM 调用报 `RPC server is unavailable`。规避：改用 `--headed`（可见窗口）运行。 |
| **空白布局无占位符** | 空白布局（layout 7/12）没有占位符，`target={"type":"title/subtitle/body"}` 匹配不到。规避：改用 `name` / `text_match` / `position` / `index` 定位，或先 `add_textbox` 再按 name 改。 |
| **大文件 inspect** | shape 数很大时 pywin32 inspect 较慢；已改为索引遍历避免 COM 枚举器失效。需要单页时 VBA backend 可用 `InspectSlideJson(slideIndex)`。 |
| **非文本 shape 改字体** | `modify_font` / `set_alignment` 对无文本框的 shape（如图片）会跳过并返回提示，不再抛异常。 |

## 适用场景

- 🔒 涉密/气隙（air-gapped）环境
- 🏢 企业内网无外网访问
- 🤖 本地 AI Agent（Claude Code、Cursor）直接编辑 PPTX
- 💰 无需 API 费用的长期使用

## 参考文档

- `references/setup-guide.md` - Ollama 安装与模型选择详细指南
- `references/interactive-session-guide.md` - 交互会话命令协议
- `references/PptEditorBridge.bas` - VBA backend 桥接模块模板
- `references/JsonConverter.bas` - PowerPoint 兼容的精简 JSON 解析/序列化模块，供 VBA backend 使用

## VBA 开发踩坑指南

在为本 skill 编写或修改 VBA 代码时，务必注意以下经过实际调试验证的陷阱。

### 1. 引号转义极易出错 — 优先用 `Chr(34)`

VBA 用 `""` 在字符串内表示一个双引号，层层嵌套后极难数清：

```vba
' ❌ 少一个引号 → VBA 编译器死循环卡死（不报错！）
text = Replace(text, """", "\"")

' ✅ 正确：5 个引号字符组成第三参数
text = Replace(text, """", "\""")

' ✅ 推荐：用 Chr(34) 彻底避免数引号
text = Replace(text, Chr(34), "\" & Chr(34))
```

> **教训：** 一个引号的 typo 导致整个 VBA 模块编译时死循环，所有 Public Function 均不可用，`On Error Resume Next` 也捕获不了编译期错误。排查耗时数小时。

### 2. 对象赋值必须用 `Set`

VBA 区分值赋值 (`Let`) 和对象引用赋值 (`Set`)。遗漏 `Set` 不一定有编译错误，但运行时会抛 **Error 450**：

```vba
' ❌ Let 赋值 — Collection/Dictionary 没有默认属性 → Error 450
dict("items") = myCollection
value = FunctionReturningObject()

' ✅ Set 赋值
Set dict("items") = myCollection
Set value = FunctionReturningObject()
```

**需要 `Set` 的典型场景：**
- `Scripting.Dictionary` 的 Item 赋值存放 Collection / Dictionary
- 函数返回值是 Object 时的变量接收
- `For Each ... In` 循环中也是对象时的变量接收

**快速自检：** 搜索 `") = ` 模式，右侧如果是 Collection、Dictionary 或返回 Object 的函数调用，必须有 `Set`。

### 3. VBA 编译错误会静默卡死

与 Python 的 `SyntaxError` 不同，VBA 编译期问题可能：

| 表现 | 原因 |
|------|------|
| COM 调用永久挂起 | VBA 弹出错误对话框等待用户点击，但 COM Automation 看不到 UI |
| 模块级卡死 | 畸形语法导致编译器循环（如上述引号 bug） |
| `On Error` 无效 | `On Error Resume Next` 只捕获运行时错误，不捕获编译期错误 |

**建议：**
- 每个 VBA 模块都提供一个零依赖的 `Ping()` 函数用于健康检查
- 新代码先单独测试编译，再与其他模块集成
- 复杂模块做增量添加，每加一个函数就测一次

### 4. 第三方 VBA 库不能直接用于 PowerPoint

社区 VBA 库（如 VBA-JSON v2.3.1）通常为 **Excel** 开发和测试。直接导入 PowerPoint 可能因以下内容导致编译卡死：

| 不兼容内容 | 示例 |
|-----------|------|
| Windows API 声明 | `Private Declare PtrSafe Function ... Lib "kernel32"` |
| 自定义 Type 结构 | `Private Type utc_SYSTEMTIME` |
| Mac 条件编译 | `#If Mac Then ... Lib "/usr/lib/libc.dylib"` |

**建议：** 不从外部下载覆盖 skill 内置的 `JsonConverter.bas`，它是专门为 PowerPoint VBA 裁剪的精简版。

### 5. JSON 解析要正确处理空格

Python 的 `json.dumps()` 默认在 `:` 和 `,` 后加空格：

```python
json.dumps({"key": "val"})  # → {"key": "val"}  注意冒号后有空格
```

VBA 侧的 JSON 解析器在读取 value 前必须跳过空格，否则 `"target": {` 中空格后的 `{` 无法被识别为对象起始符，导致用 `Let` 而非 `Set` 赋值，触发 Error 450。
