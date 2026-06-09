# exec-csharp 云端部署方案

> **版本**：v1.0
> **日期**：2026-06-09
> **作者**：悟空 🐒
> **定位**：在 [exec-csharp 生产化方案](exec-csharp-production-plan.md) 基础上，将本地架构搬到云端 Windows VM。
> **核心决策**：保留 exec-csharp（CodeAct）作为唯一执行策略，不引入 python-pptx 等替代方案。

---

## 1. 架构总览

### 1.1 本地 vs 云端

```
本地架构（单机）：
  Agent → Python 网关 → stdin/stdout → PptInteropHost.exe → COM → Office
  （全在一台 Windows 机器上）

云端架构（分布式）：
  Agent (任意平台) → HTTP API → 云端 Windows VM → PptInteropHost.exe → COM → Office
  （Agent 在 Linux，PPT 编辑在 Windows VM）
```

### 1.2 云端架构图

```
┌────────────────────────────────────────────────────────────────┐
│                      用户 / Agent 层                            │
│  ┌──────────┐  ┌────────────────┐  ┌────────────────────┐      │
│  │ Web 前端  │  │ Agent (LLM)    │  │ API Client (SDK)   │      │
│  └────┬─────┘  └───────┬────────┘  └─────────┬──────────┘      │
│       └────────────────┼──────────────────────┘                 │
│                        │ HTTPS                                  │
└────────────────────────┼────────────────────────────────────────┘
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│                  PPT 编辑服务 (API Gateway)                      │
│                  Linux 容器 / Azure Functions                    │
│                                                                 │
│  ┌─────────────┐  ┌──────────────┐  ┌──────────────────────┐   │
│  │ REST API     │  │ 文件管理      │  │ VM 池管理             │   │
│  │ /edit        │  │ Azure Blob   │  │ 预热/分配/回收/缩容   │   │
│  │ /inspect     │  │              │  │                      │   │
│  │ /status      │  │              │  │ ┌────────────────┐   │   │
│  │ /health      │  │              │  │ │ VM 状态机       │   │   │
│  └──────┬──────┘  └──────┬───────┘  │ │ idle→busy→cool │   │   │
│         │                │          │ └────────────────┘   │   │
│         │                │          └──────────┬───────────┘   │
└─────────┼────────────────┼─────────────────────┼───────────────┘
          │                │                     │
          ▼                ▼                     ▼
┌─────────────────────────────────────────────────────────────────┐
│              Windows VM 池 (Worker)                              │
│                                                                 │
│  ┌── VM-1 (busy) ────────────────────────────────────────────┐  │
│  │  Worker Agent (Python)                                    │  │
│  │    ├── gRPC 端口 50051                                    │  │
│  │    ├── 安全六层防护                                        │  │
│  │    └── PptInteropHost.exe → Roslyn → COM → PowerPoint     │  │
│  └───────────────────────────────────────────────────────────┘  │
│                                                                 │
│  ┌── VM-2 (idle/预热) ───────────────────────────────────────┐  │
│  │  Office 已启动，Host 已加载，等待分配                       │  │
│  └───────────────────────────────────────────────────────────┘  │
│                                                                 │
│  ┌── VM-3 (cooldown) ────────────────────────────────────────┐  │
│  │  5 分钟无任务，即将休眠                                     │  │
│  └───────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
```

---

## 2. 通信协议：stdin → 远程 RPC

### 2.1 本地 vs 云端通信

```
本地：  Python ──stdin──► Host.exe ──stdout──► Python   （~0ms）
云端：  Gateway ──gRPC──► Worker ──stdin──► Host.exe     （~5-20ms）
```

### 2.2 Worker Agent（VM 内运行）

```python
# worker_agent.py — VM 内 Python 服务

class PptWorkerServicer(ppt_pb2_grpc.PptWorkerServicer):

    def __init__(self):
        self.backend = CSharpBackend(visible=False)
        self._current_file = None

    async def Open(self, request, context):
        local_path = await self._download_from_blob(request.file_url)
        self.backend.open(local_path)
        self._current_file = local_path
        return ppt_pb2.OpenResponse(ok=True)

    async def Inspect(self, request, context):
        structure = self.backend.inspect()
        return ppt_pb2.InspectResponse(
            ok=True, structure_json=json.dumps(structure, ensure_ascii=False))

    async def ExecuteCodeAct(self, request, context):
        result = self.backend.safe_code_act(request.code, self._current_file)
        return ppt_pb2.ExecuteResponse(
            ok=result["ok"],
            output=result.get("output", ""),
            error=result.get("error", ""),
            rolled_back=result.get("rolled_back", False))

    async def Save(self, request, context):
        self.backend.save()
        url = await self._upload_to_blob(self._current_file)
        return ppt_pb2.SaveResponse(ok=True, download_url=url)

    async def Close(self, request, context):
        self.backend.close()
        self._current_file = None
        return ppt_pb2.CloseResponse(ok=True)
```

### 2.3 gRPC Proto

```protobuf
syntax = "proto3";
package ppteditor;

service PptWorker {
    rpc Open(OpenRequest) returns (OpenResponse);
    rpc Inspect(InspectRequest) returns (InspectResponse);
    rpc ExecuteCodeAct(ExecuteRequest) returns (ExecuteResponse);
    rpc Save(SaveRequest) returns (SaveResponse);
    rpc Close(CloseRequest) returns (CloseResponse);
    rpc Health(HealthRequest) returns (HealthResponse);
}

message OpenRequest    { string file_url = 1; bool visible = 2; }
message OpenResponse   { bool ok = 1; int32 slide_count = 2; string error = 3; }
message InspectRequest {}
message InspectResponse { bool ok = 1; string structure_json = 2; string error = 3; }
message ExecuteRequest { string code = 1; }
message ExecuteResponse { bool ok = 1; string output = 2; string error = 3; bool rolled_back = 4; }
message SaveRequest    { string output_path = 1; }
message SaveResponse   { bool ok = 1; string download_url = 2; string error = 3; }
message CloseRequest   {}
message CloseResponse  { bool ok = 1; }
message HealthRequest  {}
message HealthResponse { bool healthy = 1; int64 uptime_seconds = 2; }
```

---

## 3. VM 生命周期管理

### 3.1 VM 状态机

```
              ┌──────────┐
 创建 VM ────►│ starting │
              └────┬─────┘
                   │ Office + Host 就绪
                   ▼
              ┌──────────┐  分配任务  ┌──────────┐
              │   idle   │──────────►│   busy   │
              │ (预热池) │           │ (工作中) │
              └────┬─────┘◄──────────└──────────┘
                   │       任务完成
              5min 无任务
                   ▼
              ┌──────────┐
              │ cooldown │─── 新任务 ──► busy
              └────┬─────┘
              超时 │
                   ▼
              ┌──────────┐
              │ stopped  │  (休眠/销毁)
              └──────────┘
```

### 3.2 池管理策略

```python
class VMPoolManager:
    MIN_IDLE = 0          # 省钱模式：不预热
    MAX_IDLE = 2          # 最大预热数
    MAX_TOTAL = 5         # 总 VM 上限
    COOLDOWN_SEC = 300    # 空闲 5 分钟 → cooldown
    STOP_AFTER_SEC = 600  # cooldown 10 分钟 → 休眠

    async def acquire_vm(self) -> VMInstance:
        # 优先级：idle > cooldown > 创建新 VM > 排队
        idle = self._find_idle_vm()
        if idle: idle.state = "busy"; return idle

        cooling = self._find_cooldown_vm()
        if cooling: cooling.state = "busy"; return cooling

        if self._total_count() < self.MAX_TOTAL:
            vm = await self._create_vm()
            vm.state = "busy"; return vm

        raise VMPoolExhausted("所有 VM 都在忙")

    async def release_vm(self, vm):
        vm.state = "idle"; vm.last_active = time.time()
        # 超上限就关最老的
        if self._idle_count() > self.MAX_IDLE:
            await self._stop_vm(self._oldest_idle())
```

### 3.3 冷启动优化

| 阶段 | 耗时 | 优化 |
|------|------|------|
| VM 启动 | ~30-60s | 预热池 0-2 台 idle |
| Office 启动 | ~3-5s | 镜像预装 + Host 开机自启 |
| Roslyn 首次编译 | ~200ms | 启动时预热空脚本 |
| PPT 打开 | ~1-2s | 依赖文件大小，无法优化 |
| **总冷启动** | **~35-68s** | **预热后：~4-7s** |

### 3.4 VM 镜像（Golden Image）

```powershell
# 1. Office 365 静默安装
Start-Process -Wait "setup.exe" "/configure office-config.xml"
# 2. .NET Runtime
winget install Microsoft.DotNet.Runtime.9
# 3. 部署 Host + Worker
Copy-Item "\\share\PptInteropHost\*" "C:\ppt-editor\" -Recurse
pip install -r C:\ppt-editor\requirements.txt
# 4. 开机自启
New-Service -Name "PptWorker" -BinaryPathName "python C:\ppt-editor\worker_agent.py" -StartupType Automatic
# 5. 预热 COM + Roslyn
$ppt = New-Object -ComObject PowerPoint.Application; $ppt.Quit()
& "C:\ppt-editor\PptInteropHost.exe" --warmup
# 6. 封装
Stop-Computer
```

---

## 4. 文件流转

### 4.1 流程

```
上传: User ──HTTPS──► Blob Storage
修改: Blob ──下载──► VM ──COM修改──► VM ──上传──► Blob
下载: Blob ──SAS URL──► User
```

### 4.2 Blob 存储结构

```
ppt-editor-files/
├── uploads/{session_id}/{upload_id}.pptx     # 原始文件
├── snapshots/{session_id}/{timestamp}.pptx   # 修改前快照
└── results/{session_id}/{edit_id}.pptx       # 修改后结果
```

### 4.3 生命周期

| 类型 | 保留 | 清理 |
|------|------|------|
| 原始文件 | 24h | Lifecycle Policy 自动删 |
| 快照 | 1h | 成功后立即删；失败保留 1h |
| 结果文件 | 24h | 下载后标记，24h 删 |

---

## 5. API 设计

### 5.1 REST API

```
POST   /api/v1/sessions                 创建编辑会话
GET    /api/v1/sessions/{id}/inspect    PPT 结构
POST   /api/v1/sessions/{id}/edit       执行 C# 脚本
POST   /api/v1/sessions/{id}/save       保存 + 下载链接
DELETE /api/v1/sessions/{id}            关闭会话
GET    /api/v1/health                   健康检查
```

### 5.2 编辑请求/响应

```json
// POST /api/v1/sessions/{id}/edit
// Request:
{ "code": "SetFont(Title(1), bold: true, colorBgr: 0xFF);", "timeout_seconds": 30 }

// Response (成功):
{ "ok": true, "output": "done", "elapsed_ms": 156, "stage": "success" }

// Response (失败):
{ "ok": false, "error": "编译失败: ...", "stage": "compilation", "rolled_back": true }
```

### 5.3 Agent 典型调用流

```python
async def edit_ppt(file_url: str, instruction: str):
    # 1. 创建会话
    s = await client.post(f"{base}/sessions", json={"file_url": file_url})
    sid = s.json()["session_id"]

    # 2. 获取结构
    structure = (await client.get(f"{base}/sessions/{sid}/inspect")).json()

    # 3. LLM 生成 C# 脚本
    script = await llm.generate_csharp(instruction, structure)

    # 4. 执行（失败则重试一次）
    result = await client.post(f"{base}/sessions/{sid}/edit", json={"code": script})
    if not result.json()["ok"]:
        script = await llm.regenerate(result.json()["error"])
        result = await client.post(f"{base}/sessions/{sid}/edit", json={"code": script})

    # 5. 保存 + 下载
    saved = await client.post(f"{base}/sessions/{sid}/save")
    return saved.json()["download_url"]
```

---

## 6. 安全：十层防护

### 6.1 原六层 → 云端映射

| 层 | 本地 | 云端 | 变化 |
|---|------|------|------|
| L1 静态预检 | Python 侧 | API Gateway 侧 | ⬆ 前置到 Gateway |
| L2 Roslyn 沙箱 | Host 内 | Host 内（VM） | ✅ 不变 |
| L3 快照回滚 | 本地文件 | Blob + 本地 | ⬆ 双重快照 |
| L4 执行超时 | CancellationToken | CT + API 超时 | ⬆ 双重超时 |
| L5 输出限制 | 1MB | 1MB + API 限制 | ⬆ 双重限制 |
| L6 异常隔离 | try/catch | try/catch + VM | ⬆ 进程 + VM 级 |

### 6.2 云端新增四层

| 层 | 防护 | 作用 |
|---|------|------|
| **L7 VM 隔离** | 每会话独占 VM | 用户间完全隔离 |
| **L8 网络隔离** | NSG：仅 gRPC 入 + Blob 出 | VM 无法访问互联网 |
| **L9 文件隔离** | 独立临时目录 | 脚本不能读其他会话文件 |
| **L10 身份认证** | API Key + JWT | 防未授权访问 |

### 6.3 NSG 规则

```
入站：✅ 50051/tcp（仅 Gateway 子网）  ❌ 其他全拒
出站：✅ 443 → Blob + KMS             ❌ 其他全拒（含互联网）
```

---

## 7. 成本分析

### 7.1 VM 定价（Azure 东亚按需）

| 规格 | 配置 | 月费 | 场景 |
|------|------|:---:|------|
| Standard_B2s | 2C 4GB | ~$38 | 单用户低频 |
| Standard_B4ms | 4C 16GB | ~$150 | 多会话并发 |
| Standard_D4s_v5 | 4C 16GB | ~$185 | 大文件高性能 |

### 7.2 省钱策略

| 策略 | 节省 |
|------|:---:|
| 按需启停（无任务休眠） | 60-80% |
| Spot 实例 | 60-90% |
| Reserved 1 年 | 40-60% |
| 预热池=0（接受冷启动） | 最大 |
| 夜间/周末自动关机 | 按需 |

### 7.3 典型月费

| 场景 | 日均编辑 | 月费 |
|------|:---:|:---:|
| 个人/Demo | ~10 | ~$5 |
| 小团队 | ~50 | ~$20 |
| 中等负载 | ~200 | ~$38-150 |
| 高负载 | ~1000+ | ~$100-450 |

+ Blob ~$2 + Office 365 ~$12.5/用户 + API Gateway ~$0-5

---

## 8. 部署方案

### 8.1 MVP（最小可行）

单台 VM，FastAPI 直接暴露 HTTP，无 VM 池：

```
Agent ──HTTP──► Windows VM
                ├── FastAPI (8080)
                ├── PptInteropHost.exe
                └── Office (预装)
```

```python
# mvp_server.py

from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import FileResponse
from ppt_backend import CSharpBackend, precheck_script, SnapshotManager
import tempfile, shutil

app = FastAPI(title="PPT Editor Cloud")
backend = CSharpBackend(visible=False)

@app.post("/api/v1/edit")
async def edit(file: UploadFile, code: str):
    ok, reason = precheck_script(code)
    if not ok:
        raise HTTPException(400, f"预检失败: {reason}")

    tmp = tempfile.mktemp(suffix=".pptx")
    with open(tmp, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        backend.open(tmp)
        snapshot = SnapshotManager(tmp)
        snapshot.take_snapshot()

        result = backend.safe_code_act(code, tmp)
        if result["ok"]:
            backend.save()
            snapshot.cleanup()
            return FileResponse(tmp, filename="edited.pptx")
        else:
            snapshot.rollback()
            raise HTTPException(422, result)
    finally:
        backend.close()
```

### 8.2 生产部署

```
Phase 1: MVP 单 VM（1-2 周）
  └── FastAPI + Host + Office，验证 e2e 流程

Phase 2: API Gateway + VM 池（2-4 周）
  └── 分离 Gateway 和 Worker，加池管理

Phase 3: 自动扩缩 + 监控（4-6 周）
  └── 按负载自动创建/销毁 VM，接入 Azure Monitor
```

---

## 9. 与 LandGod 集成

### 9.1 现有能力复用

当前已有 LandGod Worker 运行在阿里云 Windows VM 上（`iZnx3gli2zki2xZ`），可以直接复用：

```
方案 A：LandGod Worker 作为 PPT 编辑后端
  Agent → LandGod Gateway → Windows Worker → exec PptInteropHost
  优势：零新增基础设施，复用现有调度
  劣势：LandGod 协议是命令执行，不是长连接会话

方案 B：LandGod 负责 VM 生命周期，独立 Worker Agent
  LandGod 启停 VM + 部署 Worker Agent
  Agent 直连 Worker Agent gRPC 端口
  优势：会话式长连接，适合多步编辑
  劣势：需要额外网络配置
```

### 9.2 推荐：方案 B

LandGod 管 VM 生命周期（启动/停止/健康检查），Worker Agent 管 PPT 编辑会话。

```
LandGod Gateway                    Worker Agent (gRPC)
     │                                    │
     ├── 启动 VM                          │
     ├── 部署 Worker                      │
     ├── 健康检查                    ◄─────┤ /health
     ├── 停止 VM                          │
     │                                    │
Agent ────────────────────────────────────►│
     │  Open / Inspect / Edit / Save      │
```

---

## 10. 与本地方案的关系

### 10.1 代码复用率

| 组件 | 本地 | 云端 | 复用 |
|------|------|------|:---:|
| PptInteropHost.exe | ✅ | ✅ | 100% |
| PptApi.cs | ✅ | ✅ | 100% |
| precheck_script() | ✅ | ✅ | 100% |
| SnapshotManager | ✅ | ✅ | 100% |
| safe_code_act() | ✅ | ✅ | 100% |
| 审计日志 | ✅ | ✅ | 100% |
| ppt_backend.py | ✅ | ✅（VM 内） | 95% |
| pptx_editor_llm.py | ✅ | 改造为 API Client | 50% |
| VM 池管理 | ❌ | ✅ | 新增 |
| gRPC Worker | ❌ | ✅ | 新增 |
| Blob 文件管理 | ❌ | ✅ | 新增 |

**核心执行层 100% 复用**，只需要新增网络层 + VM 管理层。

### 10.2 部署模式选择

| 场景 | 推荐 |
|------|------|
| 开发/调试 | 本地（直连 COM） |
| 个人使用 | 本地 or MVP 单 VM |
| SaaS 服务 | 云端 VM 池 |
| 企业内部 | 本地 + 内网 VM |

---

## 11. 总结

> **云端 exec-csharp = 本地方案 + 网络层 + VM 生命周期**。
>
> 核心执行引擎（PptInteropHost + 六层防护）100% 复用，不需要改一行。
> 新增的只是：
> 1. **远程通信**：gRPC Worker Agent
> 2. **文件流转**：Azure Blob 上传/下载
> 3. **VM 管理**：池化 + 自动扩缩 + 按需启停
> 4. **安全增强**：VM/网络/文件隔离 + 身份认证（L7-L10）
>
> MVP 一台 VM + FastAPI 就能跑，生产环境加 VM 池 + Gateway 就完整了。
>
> **最贵的不是技术，是 Office 授权** 🐒
