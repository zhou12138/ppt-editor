# 方案三：本地 LLM + 本地 PPTX（完全离线）

## 方案概述

LLM 和 PowerPoint 编辑全部在同一台 Windows 机器上完成，使用本地 LLM（Ollama / LM Studio）替代云端 API。**零云端依赖，完全离线运行**。

适用于：涉密文档、内网环境、无互联网的气隙网络。

---

## 架构图

```
┌──────────────────────────────────────────────────────────────┐
│                     本地 Windows 机器                         │
│                                                              │
│  用户输入（命令行/交互模式）                                   │
│       │                                                      │
│       ▼                                                      │
│  pptx_editor_llm.py                                         │
│       │                                                      │
│       ├──→ 本地 LLM 服务 (localhost:11434)                   │
│       │    ┌─────────────────────┐                           │
│       │    │ Ollama / LM Studio  │                           │
│       │    │ Qwen2.5-7B / 14B   │                           │
│       │    └─────────────────────┘                           │
│       │         │ 返回 JSON 编辑指令                          │
│       │         ▼                                            │
│       ▼                                                      │
│  pptx_editor_com.py                                         │
│       │                                                      │
│       ▼                                                      │
│  COM 接口 ──→ PowerPoint.exe ──→ 编辑 .pptx                 │
│                                                              │
│  🔒 所有数据留在本地，不出网                                  │
└──────────────────────────────────────────────────────────────┘
```

---

## 前置条件

1. **硬件要求**
   - CPU: 8 核以上（推荐）
   - 内存: 16GB+（7B 模型需 ~6GB，14B 需 ~12GB）
   - GPU（可选）: NVIDIA 显卡 + CUDA 可大幅加速推理
   - 磁盘: 模型文件约 4-10GB

2. **软件要求**
   - Windows 10/11
   - Microsoft Office（含 PowerPoint）
   - Python 3.8+ + `pywin32`
   - 本地 LLM 服务（二选一）：
     - [Ollama](https://ollama.com/) —— 推荐，简单易用
     - [LM Studio](https://lmstudio.ai/) —— GUI 界面

3. **推荐模型**
   - **Qwen2.5-7B-Instruct** —— 最低要求，中文能力好
   - **Qwen2.5-14B-Instruct** —— 推荐，意图解析更准确
   - **DeepSeek-V2-Lite** —— 备选
   - ❌ 不推荐 <7B 的模型，JSON 输出不够稳定

---

## 部署步骤

### 1. 安装 Ollama

```powershell
# 下载安装 Ollama: https://ollama.com/download/windows
# 安装完成后拉取模型
ollama pull qwen2.5:7b
# 或更大的模型（效果更好）
ollama pull qwen2.5:14b
```

### 2. 验证 LLM 服务

```powershell
# Ollama 默认运行在 localhost:11434
curl http://localhost:11434/v1/chat/completions -H "Content-Type: application/json" -d "{\"model\":\"qwen2.5:7b\",\"messages\":[{\"role\":\"user\",\"content\":\"你好\"}]}"
```

### 3. 安装 Python 依赖

```powershell
pip install pywin32 openai
```

### 4. 配置环境变量

```powershell
# 指向本地 Ollama
set OPENAI_API_BASE=http://localhost:11434/v1
set OPENAI_API_KEY=ollama
set OPENAI_MODEL=qwen2.5:7b
```

如使用 LM Studio：

```powershell
set OPENAI_API_BASE=http://localhost:1234/v1
set OPENAI_API_KEY=lm-studio
set OPENAI_MODEL=qwen2.5-7b-instruct
```

### 5. 部署项目文件

```powershell
# 将以下文件放到 C:\pptx-editor\
# - pptx_editor_com.py
# - pptx_editor_llm.py
```

### 6. 验证 COM 环境

```powershell
cd C:\pptx-editor
python demo_test.py
```

### 7. 开始使用

```powershell
python pptx_editor_llm.py report.pptx "把标题改成红色大字"
```

---

## 使用示例

### 基本编辑

```bash
# 修改标题
python pptx_editor_llm.py report.pptx "第一页标题改成'内部报告'，加粗，字号36"

# 批量操作
python pptx_editor_llm.py report.pptx "所有页面背景改成深蓝色，标题字体改成白色"

# 添加内容
python pptx_editor_llm.py report.pptx "在最后新增一页，标题写'谢谢'，居中显示"
```

### 交互模式（推荐）

```bash
python pptx_editor_llm.py report.pptx --interactive

> 查看所有幻灯片的内容
📄 第1页: 标题="年度报告", 副标题="2024年"
📄 第2页: 标题="目录", 内容=4个要点
📄 第3页: 标题="业绩概览", 含1个图表
...

> 第三页的图表数据改一下，Q4从85改成92
✅ 已修改图表数据

> 保存退出
✅ 已保存 report.pptx
```

### 写一个自动化脚本

```python
import os
os.environ["OPENAI_API_BASE"] = "http://localhost:11434/v1"
os.environ["OPENAI_API_KEY"] = "ollama"
os.environ["OPENAI_MODEL"] = "qwen2.5:7b"

from pptx_editor_llm import PPTXEditorLLM

# 批量处理多个文件
pptx_files = ["report1.pptx", "report2.pptx", "report3.pptx"]
for f in pptx_files:
    editor = PPTXEditorLLM(f)
    editor.execute("把公司名称从'旧公司'替换为'新公司'")
    editor.execute("页脚添加'机密文件'水印")
    editor.save()
    editor.close()
    print(f"✅ {f} 处理完成")
```

---

## 注意事项 / 已知问题

### ⚠️ 模型质量直接影响效果

本地 LLM 的意图解析能力不如 GPT-4 / Claude。常见问题：

| 问题 | 原因 | 解决方案 |
|------|------|---------|
| JSON 格式错误 | 小模型输出不稳定 | 用 14B+ 模型，或加重试逻辑 |
| 理解不准确 | 模型能力有限 | 用更明确的指令，避免模糊表达 |
| 中文支持差 | 某些模型中文训练不足 | 使用 Qwen 系列（中文最佳） |

**推荐**：至少使用 **Qwen2.5-7B**，最佳体验用 **14B**。

### ⚠️ 首次推理较慢

本地模型首次加载需要 10-30 秒（取决于模型大小和硬件），之后每次请求约 2-10 秒。有 GPU 可显著加速。

### ⚠️ 桌面会话

与方案二相同，COM 必须在桌面会话中运行。本地使用时通常不是问题（直接在桌面操作）。

### ⚠️ 内存占用

同时运行 LLM + PowerPoint 较占内存：
- Ollama (7B 模型): ~6GB
- PowerPoint: ~500MB-1GB
- 建议总内存 16GB+

---

## 文件清单

| 文件 | 说明 |
|------|------|
| `pptx_editor_com.py` | COM 编辑器核心，70+ 方法 |
| `pptx_editor_llm.py` | LLM 集成层，自然语言→编辑指令 |
| `demo_test.py` | 测试套件，验证环境 |

## 与其他方案对比

| | 方案一（LandGod） | 方案二（云端） | **方案三（本地）** |
|--|---|---|---|
| 隐私性 | 中（文件经网络） | 低（数据上云） | **高（全离线）** |
| 部署难度 | 高 | 中 | **中** |
| 成本 | LandGod + 服务器 | API 按量付费 | **一次性硬件成本** |
| 效果 | 取决于 LLM | 最佳 | **取决于模型大小** |
| 适用场景 | 跨平台团队 | 通用 | **涉密/内网** |
