# PptInteropHost

常驻 .NET 控制台进程,通过 **NDJSON over stdin/stdout** 协议接受指令,用**晚绑定 COM(`dynamic`/IDispatch)** 操控 PowerPoint。是 benchmark 里的 "C# EXE" backend,也是 CodeAct 方案的载体。

## 架构

```
调用方(Python/LLM) ──换行分隔JSON(stdin/stdout)──▶ PptInteropHost.exe ──COM(late binding)──▶ PowerPoint.exe
```

- COM 用 `dynamic` 晚绑定(与 pywin32 相同的 IDispatch 机制),零 NuGet/PIA/GUID 依赖。
- 进程常驻:一次启动多次复用,省去反复启动 PowerPoint。
- 协议刻意对齐 VBA bridge,便于 benchmark 用同一驱动对比多后端。

## 通信协议(NDJSON-RPC)

**一行 = 一条消息**,以 `\n` 分隔,每条都是单行 JSON。同步、按行分帧、顺序配对。

请求:
```json
{"cmd": "open", "path": "C:/deck.pptx", "visible": true}
```

响应:
```json
{"ok": true,  "result": <任意值>}
{"ok": false, "error": "错误信息"}
```

支持的 `cmd`:`ping`、`open`、`save`、`inspect`、`inspect_slide`、`execute_action`、`execute_code`、`set_notes`、`append_notes`、`close`/`quit`。

## 两种操作方案

host 同时支持两种编辑方式,区别在于**谁来编排多步操作、往返几次**。

### 旧方案:`execute_action`(JSON)

每个操作单独编码成一条 JSON,逐条发送,host 端 `ExecuteAction` 用大 `switch` 一条做一件事。

```json
{"cmd":"execute_action","action":"modify_text","slide":1,"target":{"type":"title"},"params":{"text":"新标题"}}
{"cmd":"execute_action","action":"modify_font","slide":1,"target":{"type":"title"},"params":{"bold":true,"color":255}}
```

### 新方案:`execute_code`(CodeAct)

整个计划写成一段 C# 脚本,一次发送,host 用 Roslyn(`Microsoft.CodeAnalysis.CSharp.Scripting`)在进程内编译执行,对着 `PptApi` 契约编程。

```json
{"cmd":"execute_code","code":"var t=Title(1); SetText(t,\"新标题\"); SetFont(t,bold:true,colorBgr:0x0000FF); AddTextbox(1,\"副标题\",60,160,400,50); Print(\"done\");"}
```

返回 `{"ok":true,"result":{"output":"<Print 输出>"}}`。

### 对照

| 维度 | 旧:execute_action (JSON) | 新:execute_code (CodeAct) |
|------|-------------------------|---------------------------|
| 粒度 | 一条 = 一个原子操作 | 一段 = 整个计划 |
| 往返次数 | N 步 = N 次 stdin/stdout 往返 | N 步 = **1 次**往返 |
| 编排者 | 调用方/LLM(model→tool→model 循环) | 脚本本身(一次执行内完成) |
| 控制流 | 无(只能线性发命令) | 有:循环/if/过滤/聚合/中间变量 |
| 取值再用 | 难(inspect→模型解析→再发下一条) | 易(`var t=Title(1)` 直接拿对象继续用) |
| 契约 | JSON action 名 + params 字段 | `PptApi` 的 C# 方法签名(带类型/默认值) |
| 错误粒度 | 每条命令单独成功/失败 | 整段脚本一个成功/失败 |
| 审批 | 可逐条审批 | 只能对整段 `execute_code` 审批 |
| 灵活度 | 受限于已实现的 action 列表 | 任意逻辑 + 经 `App`/`Prs` 直达原始 COM |

**核心收益**:CodeAct 把 `model → tool → model` 循环塌缩成一次执行,消除编排开销;在本 host 里还叠加进程内收益——往返从 N 降到 1,脚本内部对 COM 的多次访问也在 host 进程内连续完成。

**何时仍用旧 JSON**:只有 1-2 步、每步需单独可见/审批、副作用要逐个确认。

## PptApi 契约(CodeAct 脚本可用 API)

定义在 `PptApi.cs`。**颜色为 BGR**(红=0x0000FF,蓝=0xFF0000);**位置/大小单位 points**(72pt=1 inch);**所有索引 1-based**。

| 分类 | 成员 |
|------|------|
| 输出 | `Print(msg)`、`Output` |
| 导航 | `SlideCount`、`Slide(i)`、`Shape(slide,idx)`、`ShapeCount(slide)`、`FindByText(slide,contains)`、`Title(slide)` |
| 文本/字体 | `SetText(shp,text)`、`GetText(shp)`、`SetFont(shp, size?, bold?, italic?, colorBgr?, name?)` |
| 几何 | `Move(shp,left,top)`、`Resize(shp,width,height)` |
| 外观 | `SetFill(shp,bgr)`、`SetBorder(shp,bgr,weight?)`、`SetSlideBackground(slide,bgr)` |
| 创建 | `AddTextbox(slide,text,left,top,width,height)` |
| 备注 | `SetNotes(slide,text)` |
| 原始 COM | `App`(Application)、`Prs`(Presentation) |

## 构建与运行

```powershell
dotnet build csharp_interop/PptInteropHost -c Release
# 产物: csharp_interop/PptInteropHost/bin/Release/net9.0-windows/PptInteropHost.exe
```

驱动示例(PowerShell,用 .NET Process 显式重定向,避免管道死锁):

```powershell
$exe = (Resolve-Path "csharp_interop/PptInteropHost/bin/Release/net9.0-windows/PptInteropHost.exe").Path
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $exe
$psi.RedirectStandardInput = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$p = [System.Diagnostics.Process]::Start($psi)
$so = $p.StandardOutput.ReadToEndAsync()
@(
  (@{cmd='open'; path=(Resolve-Path "sample.pptx").Path; visible=$false} | ConvertTo-Json -Compress),
  (@{cmd='execute_code'; code='Print($"slides={SlideCount}");'} | ConvertTo-Json -Compress),
  (@{cmd='quit'} | ConvertTo-Json -Compress)
) | ForEach-Object { $p.StandardInput.WriteLine($_) }
$p.StandardInput.Close(); $p.WaitForExit(60000) | Out-Null
$so.Result
```

> 提示:PowerPoint 仅在交互式启动下才加载 COM Add-in;但本 host 直接用 `PowerPoint.Application` 自动化,不依赖 Add-in。`visible=false` 时以无窗口方式打开演示文稿。

## 性能定位

仍是**进程外 COM**:`inspect all` 与 pywin32 一样要 N 次跨进程 IPC,只是 C# 单次封送更快(~2ms vs ~27ms),把 3.72s 压到 ~276ms,但消不掉 IPC 次数瓶颈,仍输给进程内零 IPC 的 VBA/Add-in。CodeAct 减少的是**调用方↔host 的往返次数**,与"进程内 vs 进程外"是两个正交的优化维度。

## 文件

| 文件 | 说明 |
|------|------|
| `Program.cs` | 主循环、协议分发、`execute_action` 全部 action、`execute_code`、inspect、COM 生命周期 |
| `PptApi.cs` | CodeAct 脚本的 API 契约(globals) |
| `PptInteropHost.csproj` | net9.0-windows,启用 `BuiltInComInteropSupport`,引用 Roslyn scripting |
