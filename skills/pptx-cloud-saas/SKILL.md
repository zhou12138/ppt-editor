---
name: pptx-cloud-saas
description: "Edit PowerPoint on a Windows cloud server using cloud LLM API. Both LLM and PPTX are in the cloud on the same Windows machine. Use when: cloud-hosted PowerPoint editing, SaaS-style PPTX service, Windows cloud server with Office. Triggers on: cloud pptx editing, saas powerpoint, cloud pptx service."
---

# 云端 PowerPoint 编辑（SaaS 模式）

## 概述

本方案将 LLM API 调用和 PowerPoint COM 自动化集中在同一台 Windows 云服务器上。`pptx_editor_llm.py` 作为核心枢纽，调用云端 LLM（如 OpenAI/Azure）解析用户意图，再直接通过 COM 操作本地的 PowerPoint。

## 架构图

```
┌─────────────────────────────────────────────────────┐
│  Windows 云服务器                                    │
│                                                      │
│  ┌──────────────────┐    ┌────────────────────────┐ │
│  │ pptx_editor_llm  │───►│ pptx_editor_com.py     │ │
│  │ (意图解析)        │    │ COM → PowerPoint       │ │
│  └────────┬─────────┘    └────────────────────────┘ │
│           │                                          │
│           │ API 调用                                  │
└───────────┼──────────────────────────────────────────┘
            ▼
    ┌───────────────┐
    │ Cloud LLM API │
    │ (OpenAI等)    │
    └───────────────┘
```

## 部署要求

| 组件 | 要求 |
|------|------|
| 操作系统 | Windows Server 2019+ |
| Office | Microsoft Office 2016+（含 PowerPoint） |
| Python | 3.8+ |
| pywin32 | 最新版 |
| LLM API Key | OpenAI / Azure OpenAI / 其他兼容 API |

## 使用模式

### 模式一：CLI 单次命令

```bash
python pptx_editor_llm.py --file "C:\Reports\deck.pptx" --instruction "把第一页标题改成蓝色加粗"
```

### 模式二：交互模式

```bash
python pptx_editor_llm.py --file "C:\Reports\deck.pptx" --interactive
> 把第二页的图表标题改成"2024年营收分析"
> 给第三页添加一个红色文本框，内容是"机密"
> 保存退出
```

### 模式三：HTTP API 封装（SaaS 化）

可将 `pptx_editor_llm.py` 封装为 Flask/FastAPI 服务，提供 REST API：

```python
from fastapi import FastAPI
from pptx_editor_llm import process_instruction
from pptx_editor_com import PowerPointCOM

app = FastAPI()

@app.post("/api/edit")
async def edit_pptx(file_path: str, instruction: str):
    result = process_instruction(file_path, instruction)
    return {"status": "ok", "result": result}
```

```bash
curl -X POST http://your-server:8000/api/edit \
  -d '{"file_path": "C:\\deck.pptx", "instruction": "修改标题颜色为红色"}'
```

## 代码示例

### 基本使用流程

```python
from pptx_editor_com import PowerPointCOM

p = PowerPointCOM()
p.open("C:\\presentation.pptx")

# 查看结构
info = p.inspect()
print(info)

# 编辑操作
p.set_text(1, 1, "新标题内容")
p.set_font_size(1, 1, 28)
p.set_font_color(1, 1, "FF0000")  # BGR: 蓝色

# 保存并清理
p.save()
p.cleanup()
```

## 脚本文件

skill 目录下 `scripts/` 包含所有需要的脚本：

| 文件 | 用途 |
|------|------|
| `pptx_editor_com.py` | COM 引擎，70+ 方法，核心执行层 |
| `pptx_editor_llm.py` | LLM 意图解析，支持 CLI / 交互模式 / HTTP API |

两个文件都部署在同一台 Windows 云服务器上。

## 注意事项

| 问题 | 说明 |
|------|------|
| **桌面会话** | COM 需要桌面会话，云服务器需保持 RDP 连接或配置自动登录 |
| **API 费用** | 每次编辑会调用 LLM API，注意控制 token 使用量 |
| **网络延迟** | LLM API 调用有网络延迟（通常 1-5 秒），批量操作时需考虑 |
| **BGR 颜色** | COM 使用 BGR 格式，红色为 `0000FF`，蓝色为 `FF0000` |
| **并发限制** | PowerPoint COM 不支持并发操作，多用户需排队或多实例 |

## 适用场景

- 企业内部 PPTX 批量处理服务
- Web 应用后端的 PowerPoint 生成/编辑功能
- 定时任务自动更新报告内容

## 参考文档

- `references/deployment-guide.md` - 完整部署和使用指南
