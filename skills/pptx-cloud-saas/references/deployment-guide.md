# 云端 PPTX 编辑服务部署指南

## 概述

本指南详细说明如何在 Windows 云服务器上部署 PPTX 编辑服务，实现 SaaS 化的 PowerPoint 自动化编辑。

## 环境准备

### 1. Windows 服务器配置

```powershell
# 确认 Windows 版本
winver

# 确认 Office 已安装
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\*\Common\InstallRoot" -ErrorAction SilentlyContinue
```

### 2. Python 环境安装

```powershell
# 安装 Python 3.10+
winget install Python.Python.3.10

# 安装依赖
pip install pywin32 openai requests
```

### 3. 配置 LLM API 密钥

```powershell
# 设置环境变量
setx OPENAI_API_KEY "sk-your-key-here"

# 或使用 Azure OpenAI
setx AZURE_OPENAI_ENDPOINT "https://your-resource.openai.azure.com/"
setx AZURE_OPENAI_KEY "your-azure-key"
```

### 4. 部署项目文件

将以下文件复制到服务器：

```
C:\pptx-editor\
├── pptx_editor_com.py    # COM 操作核心（必须）
├── pptx_editor_llm.py    # LLM 意图解析（必须）
├── pptx_editor.py        # 规则编辑器（可选）
└── demo_test.py          # 测试脚本（可选）
```

## 桌面会话配置

### 关键：保持桌面会话

COM 自动化需要桌面会话（非 Session 0）。以下方案确保会话持续：

#### 方案 A：RDP 保持连接

```powershell
# 修改组策略，防止 RDP 断开后注销
# gpedit.msc → 计算机配置 → 管理模板 → Windows 组件 → 远程桌面服务
# 设置"断开连接后保持会话"
```

#### 方案 B：自动登录 + tscon

```powershell
# 配置自动登录
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v AutoAdminLogon /t REG_SZ /d 1 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultUserName /t REG_SZ /d "your_user" /f
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultPassword /t REG_SZ /d "your_pass" /f
```

## 验证部署

```python
# test_deployment.py
from pptx_editor_com import PowerPointCOM

p = PowerPointCOM()
p.create_presentation()
p.add_slide(1)
p.set_text(1, 1, "部署测试成功！")
p.save_as("C:\\test_output.pptx")
p.cleanup()
print("✅ 部署验证通过")
```

## HTTP API 服务化

### 使用 FastAPI 封装

```python
# server.py
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from pptx_editor_com import PowerPointCOM
import threading

app = FastAPI(title="PPTX Editor API")
lock = threading.Lock()  # COM 不支持并发

class EditRequest(BaseModel):
    file_path: str
    instructions: list[dict]

@app.post("/api/edit")
async def edit_pptx(req: EditRequest):
    with lock:
        try:
            p = PowerPointCOM()
            p.open(req.file_path)
            for inst in req.instructions:
                action = inst.get("action")
                getattr(p, action)(**{k:v for k,v in inst.items() if k != "action"})
            p.save()
            p.cleanup()
            return {"status": "success"}
        except Exception as e:
            raise HTTPException(500, str(e))
```

启动服务：

```bash
uvicorn server:app --host 0.0.0.0 --port 8000
```

## 性能优化建议

- **复用 COM 对象：** 频繁编辑时保持 `PowerPointCOM` 实例，避免反复启动 PowerPoint
- **批量操作：** 将多个编辑指令合并为一次 `open-edit-save` 周期
- **内存监控：** 长期运行时监控内存使用，定期重启 PowerPoint 进程
- **API 缓存：** 对相同意图的 LLM 解析结果做缓存，减少 API 调用

## 监控与日志

```python
import logging
logging.basicConfig(
    filename="C:\\pptx-editor\\logs\\service.log",
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s"
)
```

建议接入企业监控系统（如 Prometheus + Grafana）监控服务健康状态。
