# 本地离线 PPTX 编辑环境搭建指南

## 概述

本指南详细说明如何在 Windows 机器上搭建完全离线的 PowerPoint 自动化编辑环境，使用 Ollama 运行本地 LLM 进行意图解析。

## 第一步：安装基础环境

### Python 环境

```powershell
# 安装 Python 3.10+
winget install Python.Python.3.10

# 安装 pywin32
pip install pywin32

# 运行 pywin32 安装后脚本
python -m pywin32_postinstall -install
```

### 确认 Office 可用

```python
import win32com.client
app = win32com.client.Dispatch("PowerPoint.Application")
print(f"PowerPoint 版本: {app.Version}")
app.Quit()
```

## 第二步：安装 Ollama

### Windows 安装

1. 访问 https://ollama.ai/download 下载 Windows 版本
2. 或使用 winget：

```powershell
winget install Ollama.Ollama
```

3. 安装完成后验证：

```powershell
ollama --version
```

### 启动 Ollama 服务

```powershell
# Ollama 安装后会自动注册为服务
# 手动启动（如未自动启动）
ollama serve
```

验证服务运行：

```bash
curl http://localhost:11434/api/tags
```

## 第三步：选择并下载模型

### 推荐模型：Qwen2.5 7B

Qwen2.5 系列对中文支持优秀，JSON 结构化输出稳定：

```bash
# 下载 7B 模型（约 4.5GB）
ollama pull qwen2.5:7b
```

### 模型选择建议

#### 内存 8GB 的机器

```bash
ollama pull qwen2.5:7b     # 首选
# 注意：运行 7B 模型 + PowerPoint 可能内存紧张
# 编辑前关闭不必要的应用程序
```

#### 内存 16GB+ 的机器

```bash
ollama pull qwen2.5:14b    # 质量更高，推荐
```

#### 内存仅 4GB

```bash
ollama pull qwen2.5:3b     # 勉强可用，JSON 输出可能不稳定
```

### 验证模型可用

```bash
ollama run qwen2.5:7b "请将以下指令转换为JSON格式：把标题改成红色"
```

期望输出类似：

```json
{"action": "set_font_color", "slide": 1, "shape_index": 1, "color": "0000FF"}
```

## 第四步：配置环境变量

```powershell
# 永久设置（重启后生效）
setx OPENAI_API_BASE "http://localhost:11434/v1"
setx OPENAI_API_KEY "ollama"

# 当前会话临时设置
set OPENAI_API_BASE=http://localhost:11434/v1
set OPENAI_API_KEY=ollama
```

## 第五步：部署项目文件

```
C:\pptx-editor\
├── pptx_editor_com.py    # COM 操作核心
├── pptx_editor_llm.py    # LLM 意图解析
└── pptx_editor.py        # 规则编辑器（备用）
```

## 第六步：完整测试

```python
# full_test.py - 完整离线编辑测试
import os
os.environ["OPENAI_API_BASE"] = "http://localhost:11434/v1"
os.environ["OPENAI_API_KEY"] = "ollama"

from pptx_editor_com import PowerPointCOM

# 创建测试文件
p = PowerPointCOM()
p.create_presentation()
p.add_slide(1)
p.set_text(1, 1, "测试标题")
p.save_as("C:\\test_offline.pptx")
p.cleanup()

print("✅ 离线编辑测试通过！")
```

## 常见问题

### Ollama 服务无法启动

```powershell
# 检查端口是否被占用
netstat -ano | findstr 11434

# 手动启动并查看日志
ollama serve 2>&1
```

### 模型推理速度慢

- 检查是否有 GPU 可用：`ollama ps` 查看运行状态
- NVIDIA GPU 用户确认已安装 CUDA 驱动
- CPU 模式下 7B 模型约 5-15 token/s，耐心等待

### JSON 输出格式错误

- 升级到更大的模型（7B → 14B）
- 在 system prompt 中明确要求 JSON 格式
- 添加 few-shot 示例提高准确率

### 内存不足

```powershell
# 关闭不必要进程
taskkill /f /im "不需要的程序.exe"

# 监控内存使用
tasklist /fi "imagename eq ollama*" /fo table
tasklist /fi "imagename eq POWERPNT.EXE" /fo table
```

## LM Studio 替代方案

如果更喜欢 GUI 界面，可以使用 LM Studio：

1. 下载 LM Studio：https://lmstudio.ai/
2. 在界面中搜索并下载 Qwen2.5 7B GGUF 模型
3. 启动本地服务器（默认端口 1234）
4. 修改环境变量：

```powershell
setx OPENAI_API_BASE "http://localhost:1234/v1"
```

LM Studio 提供模型管理 GUI，适合不熟悉命令行的用户。

## VBA Backend 配置

VBA backend (`--backend vba`) 是性能最优的通用方案，inspect 比 pywin32 快 20 倍+。首次使用时 skill 会自动导入 VBA 模块，但需要先启用 VBA 项目访问权限。

### 第一步：启用 VBA 项目信任

1. 打开 PowerPoint → 文件 → 选项 → 信任中心 → 信任中心设置
2. 左侧选择 **宏设置**
3. 勾选 ✅ **"Trust access to the VBA project object model"**（信任对 VBA 项目对象模型的访问）
4. 点击确定

> ⚠️ 如果不启用此选项，`_ensure_vba_modules()` 自动导入会失败，报 "Programmatic access to VBA project is not trusted"。

### 第二步：验证 VBA backend

```powershell
python pptx_editor_llm.py your_file.pptx --backend vba --inspect --headed
```

预期输出：
```
📥 自动导入 VBA 模块: JsonConverter
📥 自动导入 VBA 模块: PptEditorBridge
📂 已打开: your_file.pptx (N页)

==================================================
📄 第 1 页 (Title Slide)
==================================================
  [1] [2] Title 1 [CENTER_TITLE] (中中) → ...
```

### 自动导入机制

VBA backend 在 `open()` 时会自动检查并导入以下模块：

| 模块 | 来源 | 用途 |
|------|------|------|
| `JsonConverter` | `references/JsonConverter.bas` | JSON 序列化/反序列化 |
| `PptEditorBridge` | `references/PptEditorBridge.bas` | VBA 桥接（inspect/execute） |

如果目标 .pptx 已经包含这些模块（如之前导入过），不会重复导入。

### 交互式使用

```powershell
python pptx_editor_llm.py your_file.pptx --backend vba --interactive-actions --headed
```

然后输入 JSON actions：
```json
{"action":"modify_font","slide":1,"target":{"type":"title"},"params":{"bold":true,"color":65280}}
```

### 常见问题

#### "Sub or function not defined"

**原因**：VBA 模块未导入（Trust Center 未启用或文件只读）  
**解决**：检查 Trust Center 设置，确保文件不是只读模式

#### "Wrong number of arguments or invalid property assignment" (Error 450)

**原因**：PptEditorBridge 中 `Set` 关键字遗漏或 JsonConverter 版本不兼容  
**解决**：确保使用 skill 内置的 `JsonConverter.bas`，不要从外部下载通用版 VBA-JSON

#### VBA 模块导入后 inspect 挂起

**原因**：从 GitHub 下载的通用版 VBA-JSON (1122 行) 在 PowerPoint VBA 中编译死循环  
**解决**：skill 内置的是专门为 PowerPoint 裁剪的精简版 (266 行)，不要替换

#### 中文内容匹配失败

**原因**：VBA `InStr + vbTextCompare` 对中文字符不可靠  
**解决**：使用 `target: {"name": "..."}` 或 `target: {"index": N}` 替代 `text_match`
