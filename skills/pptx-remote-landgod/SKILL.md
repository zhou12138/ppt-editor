---
name: pptx-remote-landgod
description: "Edit PowerPoint files on a remote Windows machine via LandGod Gateway. LLM runs locally (Linux/Mac), sends COM automation commands to Windows Worker through LandGod API. Use when: user wants to edit PPTX on a remote Windows machine, remote PowerPoint automation, cross-platform PPTX editing. Triggers on: edit pptx remotely, landgod pptx, remote powerpoint."
---

# 远程 PowerPoint 编辑（LandGod 网关模式）

## 概述

本方案适用于 LLM 运行在本地 Linux/Mac 上，通过 LandGod Gateway 将 COM 自动化命令发送到远程 Windows Worker 执行 PowerPoint 编辑的场景。

**数据流：** LLM（本地）→ LandGod Gateway（HTTP API）→ Windows Worker → pptx_editor_com.py → COM → PowerPoint

## 架构图

```
┌─────────────────────┐         ┌──────────────────┐         ┌─────────────────────────┐
│  本地 Linux/Mac     │  HTTP   │  LandGod Gateway │  调度   │  Windows Worker          │
│                     │  API    │  (localhost:8081) │         │                          │
│  pptx_editor_llm.py ├────────►│  /api/tool_call  ├────────►│  pptx_editor_com.py      │
│  (意图解析+代码生成) │         │  Token 认证      │         │  pywin32 + COM + Office   │
│                     │◄────────┤                  │◄────────┤                          │
│  接收结果/下载文件   │  JSON   │                  │  结果   │  PowerPoint.Application   │
└─────────────────────┘         └──────────────────┘         └─────────────────────────┘
```

## 工作流程

### 第一步：通过 LandGod 获取 PPTX 结构

```bash
curl -X POST http://localhost:8081/api/tool_call \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "command": "python",
    "args": "'$(echo -n 'from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open("C:\\test.pptx"); print(p.inspect()); p.cleanup()' | base64)'"
  }'
```

### 第二步：LLM 解析用户意图

使用 `pptx_editor_llm.py` 的 system prompt 将用户自然语言指令转换为 JSON 操作列表：

```python
# pptx_editor_llm.py 会生成类似如下的 JSON actions
actions = [
    {"action": "set_text", "slide": 1, "shape_index": 0, "text": "新标题"},
    {"action": "set_font_color", "slide": 1, "shape_index": 0, "color": "FF0000"}
]
```

### 第三步：生成 Python 代码

将 JSON actions 转换为调用 `pptx_editor_com.py` 方法的 Python 脚本。

### 第四步：通过 LandGod 发送执行

```bash
# 将 Python 脚本 base64 编码后发送
SCRIPT=$(cat <<'EOF' | base64
from pptx_editor_com import PowerPointCOM
p = PowerPointCOM()
p.open("C:\\test.pptx")
p.set_text(1, 0, "新标题")
p.set_font_color(1, 0, "0000FF")  # 注意 BGR 颜色
p.save()
p.cleanup()
EOF
)

curl -X POST http://localhost:8081/api/tool_call \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -d "{\"command\": \"python\", \"args\": \"$SCRIPT\"}"
```

### 第五步：获取结果并下载文件

从返回的 JSON 中提取执行结果，如需要可下载编辑后的文件。

## 脚本文件

skill 目录下 `scripts/` 包含所有需要的脚本：

| 文件 | 部署位置 | 用途 |
|------|---------|------|
| `pptx_editor_com.py` | Windows Worker | COM 引擎，70+ 方法 |
| `pptx_editor_llm.py` | Linux Agent 端 | LLM 意图解析 + 系统提示词 |
| `demo_test.py` | Windows Worker | 验证部署：65 个测试用例 |
| `gen_test.py` | Windows Worker | 生成测试用 PPTX |

Agent 加载 skill 后可通过 `skill_view("pptx-remote-landgod", file_path="scripts/pptx_editor_com.py")` 获取脚本内容。

## 注意事项与陷阱

| 问题 | 说明 |
|------|------|
| **4000 字符限制** | LandGod 命令有 4000 字符上限，复杂操作需拆分为多次调用 |
| **Session 0 COM 限制** | Windows 服务模式（Session 0）无法执行 `Open`/`SaveAs`，必须通过 RDP 桌面会话运行 |
| **BGR 颜色** | COM 使用 BGR 格式而非 RGB，红色应为 `0000FF` 而非 `FF0000` |
| **2GB 内存清理** | 低内存机器需在 COM 操作前清理内存，避免 PowerPoint 启动失败 |
| **集合从 1 开始** | PowerPoint COM 所有集合索引从 1 开始（非 0） |

## 前置条件

- **本地端：** Python 3.8+，`pptx_editor_llm.py`，LLM API 密钥（OpenAI/其他）
- **LandGod Gateway：** 已部署，端口 8081，Token 认证配置完成
- **Windows Worker：** Microsoft Office（含 PowerPoint），pywin32，`pptx_editor_com.py` 已部署
- Worker 需通过 RDP 保持桌面会话

## 文件部署清单

| 文件 | 部署位置 | 用途 |
|------|----------|------|
| `pptx_editor_llm.py` | 本地 Linux/Mac | LLM 意图解析、代码生成 |
| `pptx_editor_com.py` | Windows Worker | 70+ COM 方法，实际操作 PowerPoint |
| `pptx_editor.py` | Windows Worker（可选） | 规则化编辑器，无需 LLM |

## 参考文档

- `references/workflow-guide.md` - 完整工作流中文指南
- `references/com-element-targeting.md` - COM 元素定位原理
- `references/com-technical-overview.md` - pywin32 COM 技术概览
