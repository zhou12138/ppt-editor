# 交互模式指南（模式 B/C）

本指南说明 `pptx-local-offline` 新增的会话式交互能力：

- `--interactive-actions [JSON]`
- `--interactive-script [SCRIPT]`

目标：在同一个 PowerPoint COM 会话中反复执行动作，避免每次都重新打开文件。

## 设计原则

原有模式保持不变：

- `--interactive` 仍然用于模式 A 的自然语言多轮对话
- `--exec-actions` 和 `--exec-script` 仍然是一-shot 执行模式
- 新增的 `--interactive-actions` / `--interactive-script` 只为模式 C / B 提供会话能力

## 何时使用

优先级：

1. 能用 JSON actions 表达时，优先 `--interactive-actions`
2. 需要复杂控制流或 JSON 还不能覆盖时，使用 `--interactive-script`

适用场景：

- 一边看 PowerPoint 窗口一边连续微调
- 连续尝试多组 JSON actions
- 反复执行多个脚本片段
- 需要 `inspect / save / saveas / quit` 之类的会话命令

## 模式 C：JSON 会话

启动：

```powershell
python pptx_editor_llm.py deck.pptx --interactive-actions --headed
```

也可以带一个初始动作文件：

```powershell
python pptx_editor_llm.py deck.pptx --interactive-actions actions.json --headed
```

### 推荐协议：JSON command envelope

```json
{"command":"actions","payload":[{"action":"set_slide_background_image","slide":1,"params":{"image_path":"bg.png"}}]}
{"command":"inspect"}
{"command":"save"}
{"command":"quit"}
```

### 兼容输入

除了 command envelope，还支持：

- 直接输入 JSON actions 数组
- 直接输入单个 action 对象

例如：

```json
[{"action":"modify_font","slide":1,"target":{"type":"title"},"params":{"bold":true}}]
```

## 模式 B：脚本会话

启动：

```powershell
python pptx_editor_llm.py deck.pptx --interactive-script --headed
```

也可以带一个初始脚本：

```powershell
python pptx_editor_llm.py deck.pptx --interactive-script edit.py --headed
```

### 推荐协议

```json
{"command":"script","path":"edit.py"}
{"command":"inspect"}
{"command":"saveas","path":"out.pptx"}
{"command":"quit"}
```

### 支持的脚本输入

- 直接输入脚本路径
- `{"command":"script","path":"edit.py"}`
- `{"command":"script_inline","code":"ppt.set_notes(1, 'hello')"}`

## 会话命令

两种交互模式都支持：

- `help`：显示帮助
- `inspect`：输出当前结构
- `status`：显示当前会话状态
- `save`：保存到当前输出路径
- `saveas <path>`：另存为指定路径
- `quit`：保存并退出
- `close`：直接关闭并退出，不再继续会话

## 与现有参数的关系

- `--dry-run`：会执行解析与命令流，但不会真正保存
- `--headed`：建议在交互模式下开启，便于实时观察修改结果
- `--notes-progress`：会继续写入备注进度
- `--note-slide`：控制备注写入页

## 示例工作流

### 示例 1：连续调背景与标题

```powershell
python pptx_editor_llm.py deck.pptx --interactive-actions --headed --notes-progress
```

输入：

```json
{"command":"actions","payload":[{"action":"set_slide_background","slide":1,"params":{"color_bgr":16777215}}]}
{"command":"actions","payload":[{"action":"modify_font","slide":1,"target":{"type":"title"},"params":{"bold":true,"color":255}}]}
inspect
save
quit
```

### 示例 2：用脚本快速实验

```powershell
python pptx_editor_llm.py deck.pptx --interactive-script --headed
```

输入：

```text
edit_step1.py
inspect
{"command":"script_inline","code":"ppt.set_notes(1, 'reviewed')"}
saveas output_v2.pptx
quit
```

## 维护建议

- 若新增 JSON action，优先同步到 `SYSTEM_PROMPT` 和 `_dispatch`
- 会话模式的控制命令尽量保持稳定，避免频繁变更 stdin 协议
- 保持 `--interactive`、`--exec-actions`、`--exec-script` 的旧行为不变