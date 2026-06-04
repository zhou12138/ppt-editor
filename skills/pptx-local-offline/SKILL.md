---
name: pptx-local-offline
description: "Edit PowerPoint completely offline using local LLM (Ollama/LM Studio) on the same Windows machine. Zero cloud dependency. Use when: air-gapped environments, privacy-sensitive documents, offline PPTX editing, local LLM PowerPoint automation. Triggers on: offline pptx, local llm powerpoint, air-gapped pptx editing."
---

# 本地离线 PowerPoint 编辑（Ollama 模式）

## 概述

本方案在同一台 Windows 机器上运行本地 LLM（Ollama/LM Studio）和 PowerPoint COM 自动化，实现完全离线的 PPTX 编辑。零云端依赖，适合涉密环境和隐私敏感文档。

## 架构图

```
┌──────────────────────────────────────────────────┐
│  Windows 本地机器                                 │
│                                                   │
│  ┌──────────────────┐    ┌─────────────────────┐ │
│  │ pptx_editor_llm  │───►│ pptx_editor_com.py  │ │
│  │ (意图解析)        │    │ COM → PowerPoint    │ │
│  └────────┬─────────┘    └─────────────────────┘ │
│           │                                       │
│           │ localhost:11434                        │
│           ▼                                       │
│  ┌──────────────────┐                             │
│  │ Ollama / LM Studio│                            │
│  │ (本地 LLM 推理)   │                            │
│  └──────────────────┘                             │
└──────────────────────────────────────────────────┘
```

## 安装配置

### 1. 安装 Ollama

```powershell
# 下载并安装 Ollama
winget install Ollama.Ollama

# 或从官网下载：https://ollama.ai/download
```

### 2. 拉取推荐模型

```bash
# 推荐：Qwen2.5 7B（中文支持好，JSON 输出稳定）
ollama pull qwen2.5:7b

# 更高质量（需 16GB+ 内存）
ollama pull qwen2.5:14b

# 轻量级（4GB 内存可用，质量较低）
ollama pull qwen2.5:3b
```

### 3. 配置环境变量

```powershell
# 设置 API 端点指向本地 Ollama
setx OPENAI_API_BASE "http://localhost:11434/v1"
setx OPENAI_API_KEY "ollama"  # Ollama 不需要真实 key，但代码可能要求非空
```

### 4. 验证 Ollama 运行

```bash
# 确认 Ollama 服务可用
curl http://localhost:11434/v1/models

# 测试推理
curl http://localhost:11434/v1/chat/completions \
  -d '{"model":"qwen2.5:7b","messages":[{"role":"user","content":"hello"}]}'
```

## 使用示例

```bash
# 设置环境后直接使用 pptx_editor_llm.py
set OPENAI_API_BASE=http://localhost:11434/v1
set OPENAI_API_KEY=ollama

python pptx_editor_llm.py --file "C:\Reports\deck.pptx" --instruction "把标题改成红色加粗"
```

### 交互模式

```bash
python pptx_editor_llm.py --file "C:\Reports\deck.pptx" --interactive
> 第一页标题改成"项目汇报"
> 第二页添加一个文本框写"机密文件"
> 保存
```

## 模型推荐

| 模型 | 内存需求 | JSON 输出质量 | 中文支持 | 推荐场景 |
|------|----------|--------------|----------|----------|
| `qwen2.5:7b` | 8GB | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | **推荐默认** |
| `qwen2.5:14b` | 16GB | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | 高质量需求 |
| `qwen2.5:3b` | 4GB | ⭐⭐⭐ | ⭐⭐⭐⭐ | 低配机器 |
| `llama3.1:8b` | 8GB | ⭐⭐⭐⭐ | ⭐⭐⭐ | 英文场景 |
| `deepseek-coder-v2:16b` | 16GB | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | 代码生成强 |

> **关键：** 需要 7B 及以上模型才能可靠输出结构化 JSON。3B 模型可能出现格式错误。

## 注意事项

| 问题 | 说明 |
|------|------|
| **模型质量** | 本地模型的意图解析质量不如 GPT-4/Claude，复杂指令可能需要多次尝试 |
| **7B+ 要求** | 低于 7B 的模型难以可靠生成结构化 JSON 输出 |
| **推理速度** | 取决于 GPU/CPU，7B 模型在纯 CPU 上约 5-15 token/s |
| **首次加载** | 模型首次加载需 10-30 秒，后续调用复用内存中的模型 |
| **BGR 颜色** | 同其他方案，COM 使用 BGR 格式 |
| **内存竞争** | Ollama 和 PowerPoint 同时运行，建议 16GB+ 内存 |

## 适用场景

- 🔒 涉密/气隙（air-gapped）环境
- 🏢 企业内网无外网访问
- 💰 无需 API 费用的长期使用
- 🔐 隐私敏感文档处理

## 参考文档

- `references/setup-guide.md` - Ollama 安装与模型选择详细指南
