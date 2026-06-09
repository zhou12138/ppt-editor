# exec-csharp 生产化方案

> **版本**：v1.0
> **日期**：2026-06-09
> **作者**：悟空 🐒
> **前提**：基于 [C# 四策略 Benchmark 报告](../skills/pptx-local-offline/references/csharp-strategies-benchmark-report.md) 的六维度分析，选定 **exec-csharp 作为唯一生产策略**。

---

## 1. 战略决策：为什么只用 exec-csharp

| 维度 | exec-csharp | 为什么不要其他 |
|------|------------|--------------|
| **性能** | 多步操作 121ms（JSON 497ms），快 4x | 单操作慢 180ms 但人感知不到（Agent 思考 + 网络延迟已经几百ms） |
| **LLM 适配** | 代码生成是 LLM 强项，prompt 简洁 | JSON schema 生成反而是 LLM 弱项 |
| **开发成本** | 只维护 PptApi，新操作零代码 | JSON/template 每个 action 都要写代码发版 |
| **架构简洁** | 一条路径，一套监控，一个降级策略 | 四条路径 = 四套运维 = 噩梦 |

**唯一短板是安全性** → 本方案的核心就是补上这个短板。

---

## 2. 总体架构

```
┌─────────────────────────────────────────────────────────┐
│                    Agent / LLM 层                        │
│                                                         │
│  1. 接收用户指令                                         │
│  2. inspect PPT 结构                                     │
│  3. 生成完整 C# 脚本（一次性）                            │
│  4. 调用 code_act(script)                                │
└────────────────────────┬────────────────────────────────┘
                         │ C# 脚本字符串
                         ▼
┌─────────────────────────────────────────────────────────┐
│              Python 网关层 (ppt_backend.py)               │
│                                                         │
│  ┌──────────────┐  ┌──────────────┐  ┌───────────────┐  │
│  │ 脚本预检      │→ │ 快照管理      │→ │ 执行 + 超时    │  │
│  │ (静态分析)    │  │ (PPT 副本)   │  │ (CancellationToken)│
│  └──────────────┘  └──────────────┘  └───────┬───────┘  │
│                                              │          │
│  ┌──────────────┐  ┌──────────────┐          │          │
│  │ 审计日志      │← │ 结果验证      │←─────────┘          │
│  └──────────────┘  └──────────────┘                     │
└────────────────────────┬────────────────────────────────┘
                         │ stdin JSON {"cmd":"execute_code","code":"..."}
                         ▼
┌─────────────────────────────────────────────────────────┐
│           PptInteropHost.exe (常驻 C# host)              │
│                                                         │
│  ┌──────────────┐  ┌──────────────┐  ┌───────────────┐  │
│  │ Roslyn 编译   │→ │ 沙箱执行      │→ │ COM 调用       │  │
│  │ (~100-200ms) │  │ (受限命名空间) │  │ PowerPoint    │  │
│  └──────────────┘  └──────────────┘  └───────────────┘  │
│                                                         │
│  异常隔离: try/catch 全包裹，异常 → 返回错误 JSON          │
│  输出限制: Print() ≤ 1MB                                 │
└─────────────────────────────────────────────────────────┘
```

---

## 3. 安全加固（六层防护）

### 3.1 Layer 1：脚本静态预检（Python 侧）

在发送脚本到 C# host 之前，Python 侧做一轮静态文本检查：

```python
# ppt_backend.py 新增

BLOCKED_PATTERNS = [
    # 文件系统
    r'\bSystem\.IO\b',
    r'\bFile\.',
    r'\bDirectory\.',
    r'\bStreamReader\b',
    r'\bStreamWriter\b',
    # 网络
    r'\bSystem\.Net\b',
    r'\bHttpClient\b',
    r'\bWebClient\b',
    r'\bWebRequest\b',
    r'\bSocket\b',
    # 进程/反射
    r'\bSystem\.Diagnostics\b',
    r'\bProcess\.',
    r'\bAssembly\.',
    r'\bActivator\.CreateInstance\b',
    r'\bType\.GetType\b',
    # 环境
    r'\bEnvironment\.',
    r'\bRegistry\b',
    # 危险 COM
    r'\bWScript\b',
    r'\bShell\b',
    r'CreateObject\s*\(',
]

import re

def precheck_script(code: str) -> tuple[bool, str]:
    """静态检查脚本安全性。返回 (通过, 原因)。"""
    for pattern in BLOCKED_PATTERNS:
        match = re.search(pattern, code)
        if match:
            return False, f"脚本包含被禁止的 API: {match.group()}"

    # 检查脚本长度（防止 token 注入攻击）
    if len(code) > 50_000:
        return False, f"脚本过长: {len(code)} 字符（上限 50000）"

    # 检查无限循环风险
    if re.search(r'\bwhile\s*\(\s*true\s*\)', code, re.IGNORECASE):
        return False, "检测到 while(true) 无限循环"

    return True, "通过"
```

**要点**：
- 这是第一道防线，纯文本匹配，有误报可能但宁可误杀
- 被拦截后直接返回错误，不发送到 host
- 误报的合法需求通过扩展 PptApi 白名单方法解决

### 3.2 Layer 2：Roslyn 编译期沙箱（C# host 侧）

修改 `ExecuteCode` 方法，限制可用的 `using`/引用：

```csharp
// Program.cs — ExecuteCode 改造

private static object ExecuteCode(string code)
{
    if (string.IsNullOrWhiteSpace(code))
        throw new ArgumentException("缺少 code");
    if (_prs == null)
        throw new InvalidOperationException("尚未 open 演示文稿");

    // ===== 编译期沙箱：只暴露安全引用 =====
    _scriptOptions ??= ScriptOptions.Default
        .WithReferences(
            typeof(object).Assembly,                    // System.Private.CoreLib
            typeof(Enumerable).Assembly,                // System.Linq
            typeof(Microsoft.CSharp.RuntimeBinder
                .RuntimeBinderException).Assembly,      // dynamic 支持
            typeof(PptApi).Assembly)                    // PptApi
        // 只导入安全命名空间，不导入 System.IO / System.Net
        .WithImports(
            "System",
            "System.Linq",
            "System.Collections.Generic");

    var api = new PptApi(_app, _prs);

    try
    {
        using var cts = new CancellationTokenSource(
            TimeSpan.FromSeconds(EXEC_TIMEOUT_SECONDS));

        var task = CSharpScript.RunAsync(
            code, _scriptOptions, globals: api,
            cancellationToken: cts.Token);

        task.GetAwaiter().GetResult();
    }
    catch (CompilationErrorException ce)
    {
        throw new InvalidOperationException(
            "编译失败: " + string.Join(" | ", ce.Diagnostics));
    }
    catch (OperationCanceledException)
    {
        throw new InvalidOperationException(
            $"执行超时 ({EXEC_TIMEOUT_SECONDS}s)，已强制终止");
    }

    // ===== 输出限制 =====
    string output = api.Output;
    if (output.Length > MAX_OUTPUT_CHARS)
    {
        output = output.Substring(0, MAX_OUTPUT_CHARS)
            + $"\n... [截断，超出 {MAX_OUTPUT_CHARS} 字符限制]";
    }

    return new Dictionary<string, object> { ["output"] = output };
}

private const int EXEC_TIMEOUT_SECONDS = 30;
private const int MAX_OUTPUT_CHARS = 1_000_000;  // 1MB
```

**要点**：
- `ScriptOptions` 不添加 `System.IO`/`System.Net` 的 Assembly 引用 → Roslyn 编译时就会报错
- `CancellationToken` 30s 超时 → 防死循环
- 输出 1MB 上限 → 防内存爆炸

### 3.3 Layer 3：COM 状态保护（快照 + 回滚）

```python
# ppt_backend.py — 新增快照机制

import shutil
import tempfile

class SnapshotManager:
    """PPT 操作前快照，异常后回滚。"""

    def __init__(self, ppt_path: str):
        self.original_path = ppt_path
        self.snapshot_path = None

    def take_snapshot(self):
        """操作前保存副本。"""
        fd, self.snapshot_path = tempfile.mkstemp(suffix=".pptx")
        os.close(fd)
        shutil.copy2(self.original_path, self.snapshot_path)
        return self.snapshot_path

    def rollback(self):
        """异常后恢复原文件。"""
        if self.snapshot_path and os.path.exists(self.snapshot_path):
            shutil.copy2(self.snapshot_path, self.original_path)
            return True
        return False

    def cleanup(self):
        """成功后删除快照。"""
        if self.snapshot_path and os.path.exists(self.snapshot_path):
            os.remove(self.snapshot_path)
            self.snapshot_path = None
```

**Python 侧 `code_act` 封装**：

```python
def safe_code_act(self, script: str, ppt_path: str = None) -> dict:
    """带安全防护的 code_act 执行。"""

    # Layer 1: 静态预检
    ok, reason = precheck_script(script)
    if not ok:
        return {"ok": False, "error": f"脚本预检失败: {reason}", "stage": "precheck"}

    # Layer 3: 快照
    snapshot = None
    if ppt_path:
        snapshot = SnapshotManager(ppt_path)
        snapshot.take_snapshot()

    try:
        # Layer 2 + 4: 发送到 host（host 侧有编译沙箱 + 超时）
        result = self.code_act(script)

        # 成功，清理快照
        if snapshot:
            snapshot.cleanup()

        return {"ok": True, "output": result, "stage": "success"}

    except Exception as e:
        # 回滚
        if snapshot:
            rolled_back = snapshot.rollback()
            snapshot.cleanup()
        else:
            rolled_back = False

        return {
            "ok": False,
            "error": str(e),
            "stage": "execution",
            "rolled_back": rolled_back,
        }
```

### 3.4 Layer 4：执行超时（C# host 侧）

已在 Layer 2 代码中实现：
- `CancellationTokenSource(TimeSpan.FromSeconds(30))`
- 传入 `CSharpScript.RunAsync` 的 `cancellationToken`
- 超时后抛出 `OperationCanceledException`，被 catch 后返回超时错误

### 3.5 Layer 5：输出限制（C# host 侧）

已在 Layer 2 代码中实现：
- `api.Output` 截断到 1MB
- 防止 LLM 生成 `Print(巨大字符串)` 导致 host 内存爆炸

### 3.6 Layer 6：异常隔离（C# host 侧）

当前 `Program.Main()` 已有顶层 try/catch：

```csharp
try
{
    // ... Dispatch ...
    responseJson = JsonSerializer.Serialize(
        new Dictionary<string, object> { ["ok"] = true, ["result"] = result }, JsonOut);
}
catch (Exception ex)
{
    responseJson = JsonSerializer.Serialize(
        new Dictionary<string, object> { ["ok"] = false, ["error"] = FlattenMessage(ex) }, JsonOut);
}
```

**任何异常都不会杀死 host 进程**，而是返回 `{"ok": false, "error": "..."}` 并继续等待下一个请求。

---

## 4. 监控与可观测性

### 4.1 审计日志格式

每次 `code_act` 执行生成一条结构化日志：

```python
import json
import time
import logging

audit_logger = logging.getLogger("ppt_editor.audit")

def log_execution(script: str, result: dict, elapsed_ms: float):
    """记录每次脚本执行的审计日志。"""
    entry = {
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "action": "exec_csharp",
        "script_hash": hashlib.sha256(script.encode()).hexdigest()[:16],
        "script_length": len(script),
        "script_preview": script[:200],  # 前 200 字符
        "result_ok": result.get("ok", False),
        "stage": result.get("stage", "unknown"),
        "error": result.get("error"),
        "rolled_back": result.get("rolled_back", False),
        "elapsed_ms": round(elapsed_ms, 1),
        "output_length": len(result.get("output", "")),
    }
    audit_logger.info(json.dumps(entry, ensure_ascii=False))
```

### 4.2 监控指标

| 指标 | 计算方式 | 告警阈值 |
|------|---------|---------|
| **编译成功率** | 编译成功 / 总请求 | < 90% 告警 |
| **执行成功率** | 执行成功 / 编译成功 | < 95% 告警 |
| **超时率** | 超时次数 / 总请求 | > 5% 告警 |
| **回滚率** | 回滚次数 / 总请求 | > 10% 告警 |
| **预检拦截率** | 预检失败 / 总请求 | > 20% 需检查 prompt |
| **平均执行耗时** | 滑动窗口平均 | > 2000ms 告警 |
| **host 进程重启次数** | 计数器 | > 0 告警 |

### 4.3 健康检查

```python
def healthcheck(self) -> dict:
    """检查 host 进程健康状态。"""
    try:
        result = self._send_cmd({"cmd": "ping"})
        return {
            "status": "healthy",
            "host_pid": self._proc.pid,
            "uptime_seconds": time.time() - self._start_time,
        }
    except Exception as e:
        return {
            "status": "unhealthy",
            "error": str(e),
        }
```

---

## 5. 错误恢复策略

### 5.1 错误分级

| 级别 | 错误类型 | 恢复策略 |
|------|---------|---------|
| **L1 预检拦截** | 脚本包含禁止 API | 返回错误提示，让 LLM 重新生成 |
| **L2 编译失败** | Roslyn 语法/类型错误 | 返回编译诊断信息，让 LLM 修正 |
| **L3 运行时异常** | NullRef / 类型转换 / COM 异常 | 回滚快照，返回异常信息 |
| **L4 超时** | 死循环 / 长时间 COM 等待 | 强制取消，回滚快照 |
| **L5 host 崩溃** | 未捕获异常 / COM 断裂 | 重启 host 进程，重新打开 PPT |

### 5.2 LLM 重试机制

```python
MAX_RETRIES = 2

def exec_with_retry(self, script: str, ppt_path: str, model_context: dict) -> dict:
    """带 LLM 重试的执行流程。"""
    for attempt in range(1, MAX_RETRIES + 1):
        result = self.safe_code_act(script, ppt_path)

        if result["ok"]:
            return result

        if result["stage"] == "precheck":
            # 预检失败，返回给 LLM 重新生成
            return {
                "ok": False,
                "retry": False,
                "error": result["error"],
                "hint": "请修改脚本，避免使用禁止的 API",
            }

        if result["stage"] == "execution" and attempt < MAX_RETRIES:
            # 执行失败，将错误信息反馈给 LLM 重新生成脚本
            error_feedback = (
                f"脚本执行失败（第 {attempt} 次）:\n"
                f"错误: {result['error']}\n"
                f"请修正脚本并重试。"
            )
            # 这里应调用 LLM 重新生成脚本
            # script = regenerate_script(model_context, error_feedback)
            continue

        return result

    return {"ok": False, "error": "超过最大重试次数", "retry": False}
```

### 5.3 Host 进程自动重启

```python
def _ensure_host_alive(self):
    """确保 host 进程存活，崩溃后自动重启。"""
    if self._proc is None or self._proc.poll() is not None:
        exit_code = self._proc.returncode if self._proc else "N/A"
        audit_logger.warning(f"Host 进程已退出 (exit={exit_code})，正在重启...")

        self._start_host()

        # 重新打开 PPT（如果之前有打开的文件）
        if self._current_file:
            self._send_cmd({
                "cmd": "open",
                "path": self._current_file,
                "visible": self._visible,
            })

        audit_logger.info(f"Host 进程已重启 (pid={self._proc.pid})")
```

---

## 6. PptApi 扩展策略

### 6.1 当前 PptApi 覆盖的操作

```
导航: Slide(), Shape(), FindByName(), FindById(), FindByText(), Title()
      SlideCount, ShapeCount()
文本: SetText(), GetText(), SetFont()
几何: Move(), Resize()
外观: SetFill(), SetBorder(), SetSlideBackground()
创建: AddTextbox()
备注: SetNotes()
输出: Print()
原始: App, Prs (逃生舱，可访问任意 COM)
```

### 6.2 生产化需要补充的 API

| 优先级 | API | 用途 |
|--------|-----|------|
| P0 | `InspectJson()` | 脚本内获取 PPT 结构（当前只能从 Python 侧调用） |
| P0 | `Save()` / `SaveAs(path)` | 脚本内保存（Agent 完成修改后自行保存） |
| P1 | `AddSlide(index, layout)` | 添加幻灯片 |
| P1 | `DeleteSlide(index)` | 删除幻灯片 |
| P1 | `DuplicateSlide(index)` | 复制幻灯片 |
| P1 | `DeleteShape(shp)` | 删除形状 |
| P1 | `SetAlignment(shp, align)` | 设置对齐 |
| P1 | `AddPicture(slide, path, ...)` | 插入图片 |
| P2 | `AddShape(slide, type, ...)` | 添加基本形状 |
| P2 | `SetZOrder(shp, order)` | 层级调整 |
| P2 | `ModifyCell(shp, row, col, text)` | 表格单元格 |
| P2 | `SetTransition(slide, effect)` | 切换效果 |
| P2 | `ExportPdf(path)` | 导出 PDF |

### 6.3 App/Prs 逃生舱策略

当前 PptApi 暴露了 `App` 和 `Prs` 两个原始 COM 对象。这是 exec-csharp 的核心优势之一——LLM 可以直接访问任何 COM API，不受 PptApi 封装的限制。

**生产化建议**：
- 保留 `App`/`Prs` 作为逃生舱
- 通过 Python 侧静态预检控制风险（Layer 1）
- 对于频繁使用的 COM 操作，逐步封装到 PptApi 中（减少 LLM 直接操作 COM 的需求）
- 日志记录所有通过 `App`/`Prs` 直接访问的操作

---

## 7. LLM Prompt 工程

### 7.1 系统提示词（exec-csharp 专用）

```
你是一个 PowerPoint 编辑助手。你将收到 PPT 的结构信息和用户的编辑指令。

你的任务是生成一段 C# 脚本，使用 PptApi 来完成编辑。

## 可用 API

### 导航
- Slide(int slide) → 获取幻灯片对象
- Shape(int slide, int index) → 按索引获取形状
- FindByName(int slide, string name) → 按名称查找形状（返回 null 如果不存在）
- FindByText(int slide, string contains) → 按文本内容查找
- FindById(int slide, int id) → 按 ID 查找
- Title(int slide) → 获取标题占位符
- SlideCount → 幻灯片总数
- ShapeCount(int slide) → 某页形状数

### 文本 & 字体
- SetText(dynamic shp, string text)
- GetText(dynamic shp) → string
- SetFont(dynamic shp, double? size, bool? bold, bool? italic, int? colorBgr, string name)

### 位置 & 大小（单位：points，72pt = 1 inch）
- Move(dynamic shp, double left, double top)
- Resize(dynamic shp, double width, double height)

### 外观
- SetFill(dynamic shp, int colorBgr)
- SetBorder(dynamic shp, int colorBgr, double? weight)
- SetSlideBackground(int slide, int colorBgr)

### 创建
- AddTextbox(int slide, string text, double left, double top, double width, double height)

### 备注
- SetNotes(int slide, string text)

### 输出
- Print(object message) — 输出信息（类似 console.log）

### 原始 COM（高级用法）
- App → PowerPoint.Application 对象
- Prs → 当前 Presentation 对象

## 颜色规则
颜色用 BGR 整数（不是 RGB！）：
- 红色 = 0x0000FF (255)
- 蓝色 = 0xFF0000 (16711680)
- 绿色 = 0x00AA00 (43520)
- 白色 = 0xFFFFFF (16777215)
- 黑色 = 0x000000 (0)

## 脚本规则
1. 只输出 C# 代码，不要 ``` 包裹，不要解释
2. 所有 FindByName/FindByText 结果必须判空：if (shp != null) { ... }
3. 多个操作写在同一段脚本里
4. 用 Print() 输出操作结果
5. 不要使用 System.IO、System.Net 等危险 API
6. 避免 while(true) 循环
```

### 7.2 Prompt 优化要点

- **简短**：整个 prompt < 800 tokens，比 JSON schema prompt 少 60%
- **示例驱动**：附带 2-3 个 few-shot 示例
- **防御性编码**：强调 null 检查，LLM 最常犯的错就是不检查 FindByName 返回值

---

## 8. 部署 Checklist

### 8.1 C# Host 构建

```bash
# 切换到 .NET 10 （当前 csproj 写的 net9.0-windows，生产建议升级）
cd csharp_interop/PptInteropHost

# 发布为独立包（包含 Roslyn）
dotnet publish -c Release -r win-x64 --self-contained true -o ../../dist/csharp_host/

# 验证产物
ls -la ../../dist/csharp_host/PptInteropHost.exe
# 预期：~40-50MB（含 Roslyn + .NET Runtime）
```

### 8.2 目录结构

```
ppt-editor/
├── dist/
│   └── csharp_host/
│       ├── PptInteropHost.exe      # 主程序
│       ├── PptInteropHost.dll
│       ├── Microsoft.CodeAnalysis.*.dll  # Roslyn
│       └── ...
├── ppt_backend.py                  # Python 网关（含安全层）
├── pptx_editor_llm.py              # CLI 入口
├── config/
│   ├── blocked_patterns.json       # 可热更新的黑名单
│   └── exec_limits.json            # 超时/输出限制配置
└── logs/
    └── audit/                      # 审计日志
```

### 8.3 运行时依赖

| 组件 | 版本 | 必需 |
|------|------|:---:|
| Windows 11 | 10.0+ | ✅ |
| .NET Runtime | 9.0+ (或 10.0+) | ✅（self-contained 模式可不装） |
| Microsoft Office | 16.0+ (2016+) | ✅ |
| Python | 3.10+ | ✅ |
| pywin32 | 最新 | ✅（host 启动用） |

### 8.4 环境变量

```bash
# 可选：指定 host 路径（默认自动搜索）
PPTX_EDITOR_CSHARP_HOST=/path/to/PptInteropHost.exe

# 可选：超时配置
PPTX_EXEC_TIMEOUT=30

# 可选：输出限制
PPTX_MAX_OUTPUT=1000000
```

---

## 9. 测试策略

### 9.1 单元测试

| 测试类别 | 覆盖内容 | 数量 |
|---------|---------|------|
| **预检测试** | 黑名单模式匹配、长度限制、循环检测 | ~15 |
| **PptApi 测试** | 每个 API 方法的正确性 | ~25 |
| **快照测试** | 创建/回滚/清理快照 | ~5 |
| **超时测试** | 死循环脚本被正确终止 | ~3 |
| **沙箱测试** | System.IO 等被阻止 | ~10 |

### 9.2 集成测试

```python
# test_exec_csharp_e2e.py

def test_simple_font_change():
    """单操作：修改字体"""
    script = 'SetFont(Title(1), bold: true, colorBgr: 0xFF);'
    result = backend.safe_code_act(script, test_pptx)
    assert result["ok"]

def test_multi_step_user_flow():
    """多步操作：完整编辑流程"""
    script = '''
    Print(InspectJson());
    var s = FindByName(1, "TextBox 2");
    if (s != null) {
        SetFont(s, bold: true, colorBgr: 0xFF);
        SetText(s, "Modified by test");
    }
    SetNotes(1, "Test note");
    '''
    result = backend.safe_code_act(script, test_pptx)
    assert result["ok"]

def test_blocked_io():
    """沙箱：System.IO 被拦截"""
    script = 'System.IO.File.WriteAllText("hack.txt", "pwned");'
    result = backend.safe_code_act(script, test_pptx)
    assert not result["ok"]
    assert "precheck" in result["stage"]

def test_timeout():
    """超时：死循环被终止"""
    script = 'int i = 0; while (i >= 0) { i++; }'
    result = backend.safe_code_act(script, test_pptx)
    assert not result["ok"]

def test_null_safety():
    """空引用：FindByName 返回 null"""
    script = '''
    var s = FindByName(1, "NonExistentShape");
    if (s != null) SetText(s, "test");
    Print("safe");
    '''
    result = backend.safe_code_act(script, test_pptx)
    assert result["ok"]

def test_rollback_on_failure():
    """回滚：执行失败后恢复原文件"""
    import shutil
    backup = test_pptx + ".backup"
    shutil.copy2(test_pptx, backup)

    script = 'throw new System.Exception("deliberate failure");'
    result = backend.safe_code_act(script, test_pptx)
    assert not result["ok"]
    assert result.get("rolled_back")

    # 验证文件未被破坏
    assert os.path.getsize(test_pptx) == os.path.getsize(backup)
    os.remove(backup)
```

### 9.3 压力测试

```python
def test_roslyn_memory_under_load():
    """连续 100 次编译，检查内存泄漏。"""
    import psutil
    proc = psutil.Process(host_pid)
    mem_before = proc.memory_info().rss

    for i in range(100):
        script = f'Print("iter {i}"); SetNotes(1, "test {i}");'
        result = backend.safe_code_act(script, test_pptx)
        assert result["ok"]

    mem_after = proc.memory_info().rss
    growth_mb = (mem_after - mem_before) / 1024 / 1024
    assert growth_mb < 200, f"内存增长 {growth_mb:.1f}MB，可能存在泄漏"
```

---

## 10. 演进路线

### Phase 1：基础安全（本方案核心）

```
☐ Python 侧脚本静态预检 (Layer 1)
☐ C# host 编译期沙箱 (Layer 2)
☐ 执行超时 30s (Layer 4)
☐ 输出 1MB 限制 (Layer 5)
☐ 异常隔离确认 (Layer 6)——已有
☐ 审计日志
☐ 单元测试 + 集成测试
```

### Phase 2：状态保护

```
☐ PPT 快照机制 (Layer 3)
☐ Host 进程自动重启
☐ LLM 重试机制（2 次）
☐ 健康检查 endpoint
```

### Phase 3：可观测性

```
☐ 监控指标（7 个核心指标）
☐ 平滑重启（不丢失当前 PPT 状态）
☐ 压力测试（Roslyn 内存泄漏检测）
```

### Phase 4：PptApi 扩展

```
☐ P0 API：InspectJson, Save/SaveAs
☐ P1 API：AddSlide, DeleteSlide, DeleteShape, SetAlignment, AddPicture
☐ P2 API：AddShape, SetZOrder, ModifyCell, SetTransition, ExportPdf
☐ 可热更新的黑名单配置
```

---

## 11. 风险矩阵

| 风险 | 概率 | 影响 | 缓解 |
|------|:---:|:---:|------|
| LLM 生成危险代码 | 中 | 高 | Layer 1 预检 + Layer 2 沙箱 |
| Roslyn 内存泄漏 | 低 | 中 | 压测监控 + 定期重启 host |
| COM 连接断裂 | 低 | 高 | 自动重启 host + 重新打开 PPT |
| PPT 状态不可逆污染 | 中 | 高 | Layer 3 快照回滚 |
| 弱模型生成低质量代码 | 中 | 低 | LLM 重试 + 编译错误反馈 |
| 预检误报拦截合法脚本 | 低 | 低 | 扩展 PptApi 白名单方法 |

---

## 12. 总结

> exec-csharp 生产化的核心思路：**不是限制 LLM 的能力，而是给 LLM 的能力加护栏**。
>
> 六层防护层层递进：
> 1. 静态预检 → 拦截明显危险代码
> 2. 编译沙箱 → 限制可用 API
> 3. 状态快照 → 失败可回滚
> 4. 执行超时 → 防死循环
> 5. 输出限制 → 防内存爆炸
> 6. 异常隔离 → host 进程永不死
>
> **安全的本质不是信任 LLM，而是即使 LLM 犯错，系统也不会崩。**
