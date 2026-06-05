# Backend Performance Benchmark Report

> 测试日期：2026-06-05  
> 测试文件：working.pptx (4.7MB, 7 slides, 17 shapes)  
> 环境：Windows 11, Office 16 (PowerPoint 2016+), Python 3.12, .NET 9.0  
> 方法：每个 backend 独立子进程，冷启动 PPT，warmup 2 轮丢弃，测量 7 轮取均值

## 测试设计

### 公平性保证

- **完整任务计时**：所有写操作包含 find_shape + modify 的完整流程，不预缓存 shape 引用
- **进程隔离**：每个 backend 在独立子进程中运行，运行间杀掉 PPT 进程并等待 15 秒
- **pywin32 最优配置**：使用 `gencache.EnsureDispatch()` 早期绑定（比默认 `Dispatch()` 快 5x）
- **统一 warmup**：每种操作在正式测量前执行 2 轮预热
- **user_flow 模拟真实场景**：inspect → 3 次修改 → set_notes，完整用户交互流程

### 三种 Backend 架构

```
pywin32:  Python 进程 ──COM IPC (早期绑定)──→ PowerPoint 进程
VBA:      Python 进程 ──COM Application.Run──→ PowerPoint 进程内 VBA 执行
C#:       Python 进程 ──stdin/stdout 管道──→ C# EXE ──COM IPC──→ PowerPoint 进程
```

## 测试结果

### 1. 延时（完整任务）

| 任务 | pywin32 | VBA | C# | 🏅 Winner | 倍数 |
|------|---------|-----|-----|-----------|------|
| inspect 全量 (7页) | 3.91s | **235ms** | 265ms | VBA | 16.6x |
| inspect 单页 | N/A | 43ms | **38ms** | C# | 1.1x |
| modify_font (find+set) | 136ms | **29ms** | 54ms | VBA | 4.7x |
| modify_text (find+set) | 108ms | 82ms | **22ms** | C# | 4.8x |
| move_shape (find+set) | 65ms | **22ms** | 35ms | VBA | 3.0x |
| resize_shape (find+set) | 59ms | **19ms** | 32ms | VBA | 3.1x |
| set_notes (无需find) | 74ms | 63ms | **22ms** | C# | 3.3x |
| add+delete textbox | 137ms | **37ms** | 59ms | VBA | 3.7x |
| batch 10 完整任务 | 724ms | 255ms | **187ms** | C# | 3.9x |
| **user_flow** ¹ | 3.07s | **489ms** | 582ms | VBA | 6.3x |

¹ user_flow = inspect + modify_font + modify_text + move_shape + set_notes，模拟真实用户交互

### 2. 吞吐（5 秒窗口）

| 负载类型 | pywin32 | VBA | C# | 🏅 Winner |
|---------|---------|-----|-----|-----------|
| inspect 全量 | 2 (0.4/s) | **21** (4.2/s) | 20 (4.0/s) | VBA 10x |
| modify (find+set) | 43 (8.6/s) | **163** (32.6/s) | 153 (30.6/s) | VBA 4x |

### 3. 资源占用

| 指标 | pywin32 | VBA | C# | 说明 |
|------|---------|-----|-----|------|
| Host CPU | 41.0s | 0.34s | **0.28s** | pywin32 花大量 CPU 在 COM marshalling |
| PPT CPU | 68.4s | **37.2s** | 40.2s | VBA 进程内执行效率最高 |
| Host 内存 | 42.6MB | 29.4MB | **25.5MB** | C# host 最轻量 |
| PPT 内存 | 373.8MB | 396.7MB | 399.7MB | VBA/C# 多 ~25MB（模块/进程开销） |

### 4. 总分

| 排名 | Backend | 胜场 | 胜出领域 |
|------|---------|------|---------|
| 🥇 | **VBA** | **8/12** | inspect全量, modify_font, move, resize, add/del, user_flow, 两项吞吐 |
| 🥈 | **C#** | **4/12** | inspect单页, modify_text, set_notes, batch |
| 🥉 | **pywin32** | **0/12** | 无 |

## 性能差异根因分析

### 为什么 VBA 在大多数场景最快

VBA 运行在 PowerPoint 进程内部，访问 shape 属性是进程内函数调用（~0 开销），不是跨进程 COM IPC。一次 `Application.Run` 调用就能完成 find + modify 全流程，Python 端只需等待一次 IPC 返回。

```
inspect 全量：VBA 136次属性访问 × 0ms/次 = ~0ms + JSON序列化 ≈ 235ms
           pywin32 136次属性访问 × 28ms/次 = ~3800ms               ≈ 3.91s
```

### 为什么 C# 在部分场景胜出

C# 胜出的场景有两类：

**a) 不需要 find 的简单写操作 (set_notes, modify_text)**

此时瓶颈是 Python↔后端的通信开销，而非 shape 遍历：
- C# pipe 延时 ~0.3ms
- VBA Application.Run 延时 ~6ms
- 差距 20x，在简单操作中这 5.7ms 差距占比大

**b) 批量操作 (batch_10)**

10 次连续调用中，通信延时累积效应放大：
- C# 10 次 pipe = ~3ms 通信开销
- VBA 10 次 Application.Run = ~60ms 通信开销
- C# 用更快的通信弥补了 COM IPC 的劣势

**c) 单页 inspect (inspect_slide1)**

C# 有原生 `inspect_slide` 命令，只遍历 1 页的 shapes。虽然遍历仍需跨进程 COM IPC，但只有 2-3 个 shapes (~5次 IPC)，总延时很低。VBA 虽然也新增了 `InspectSlideJson`，但 Application.Run 调用本身 6ms 开销 + VBA JSON 序列化开销使得总延时略高。

### 为什么 pywin32 全面落后

即使使用了早期绑定（`EnsureDispatch`），pywin32 仍然最慢。原因：

1. **每次属性访问都是 IPC**：`shape.Name` → 跨进程 COM 调用 → 返回值 → Python 类型转换
2. **Python GIL + 动态类型**：每次 COM 调用涉及 Python 对象创建和 GC
3. **find + modify 不可合并**：必须先 N 次 IPC 找 shape，再 M 次 IPC 修改，无法一次性完成
4. **早期绑定的提升有限**：省了 `GetIDsOfNames` 查找，但 `Invoke` 本身的跨进程开销不变

### IPC 开销对比

| 通信方式 | 单次开销 | 协议 | 用于 |
|---------|---------|------|------|
| stdin/stdout 管道 | ~0.3ms | 纯文本 JSON | C# Host |
| COM Application.Run | ~6ms | COM IDispatch | VBA |
| COM 属性访问 (早期绑定) | ~10ms | COM vtable | pywin32 |
| COM 属性访问 (晚期绑定) | ~28ms | COM IDispatch | pywin32 (旧) |

### 真实用户体感

典型交互流程 `inspect → 3 次修改 → 保存`：

| Backend | 耗时 | 体感 |
|---------|------|------|
| pywin32 | 3.07s | 明显等待 |
| C# | 582ms | 略有延迟 |
| **VBA** | **489ms** | 接近即时 |

## 各 Backend 适用场景

| 场景 | 推荐 | 原因 |
|------|------|------|
| 默认 / 通用 | **VBA** | 8/12 项最快，用户体感最好 |
| 频繁全量 inspect | **VBA** | 零 IPC 遍历，17x 领先 |
| 单页交互编辑 | **C#** | 原生 `inspect_slide`，管道延时最低 |
| 批量连续操作 (>10 ops) | **C#** | 管道累积延时优势，3.9x 领先 |
| 简单单次写操作 | **C#** | 管道 0.3ms vs Application.Run 6ms |
| 零依赖 / 快速原型 | **pywin32** | 无需 VBA 宏信任或 .NET 运行时 |
| 涉密/气隙环境 | **VBA** | 只需 .bas 文件，无外部依赖 |

## 已知问题

| 问题 | 影响 Backend | 描述 |
|------|-------------|------|
| `gencache` 缓存损坏 | pywin32 | `EnsureDispatch` 首次生成的 type library 缓存偶尔损坏，需清理 `%TEMP%\gen_py` |
| COM iterator 崩溃 | pywin32 | `for shape in slide.Shapes` 在大文件 (100+ shapes) 上可能触发 RPC 断连，应改用索引遍历 |
| VBA 中文 text_match 失败 | VBA | `FindShapes` 的 `InStr + vbTextCompare` 对中文字符匹配失败 |
| VBA `target:{"index":N}` 失败 | VBA | `ParseJson` 可能将数字解析为字符串，导致 `CLng()` 类型不匹配 |
| C# `execute_action` 部分失败 | C# | 某些 action 的 JSON 字段解析不完整，返回 "未支持的 action: (空)" |
| PPT 进程残留 | 全部 | benchmark 密集运行后 PPT 进程可能无法正常退出，影响后续测试 |

## 补充：两种文件规模对比

| 指标 | working.pptx (7页/17shapes) | 柿子.pptx (8页/119shapes) |
|------|---------------------------|--------------------------|
| VBA inspect 全量 | 235ms | 946ms |
| C# inspect 全量 | 265ms | — |
| pywin32 inspect 全量 | 3.91s | 9.19s |
| VBA inspect 加速比 | 17x | 9.7x |

Shape 数量增加 7 倍后，VBA 加速比从 17x 降到 9.7x（VBA 自身的 JSON 序列化也线性增长），但仍保持数量级优势。

## Key Insights

### 1. "零 IPC" 是性能的决定性因素，语言性能几乎不重要

VBA 是世界上最慢的语言之一，但因为在 PowerPoint 进程内执行（零 IPC），它击败了 C#（JIT 编译，比 VBA 快 100 倍）和 Python 早期绑定。这颠覆了"换更快的语言就能提升性能"的直觉 — **架构（进程内 vs 跨进程）比语言选择重要得多。**

### 2. 通信开销分层，每层差距一个数量级

```
管道      0.3ms  ← C# 用这个
App.Run     6ms  ← VBA 用这个 (20x slower)
COM 调用   10ms  ← pywin32 用这个 (33x slower)
```

但 VBA 虽然通信最慢（6ms），整体却最快 — 因为它只需**通信 1 次**。pywin32 通信虽比 App.Run 只慢 1.7x，但要通信 **136 次**。**调用次数 × 单次开销 = 总开销，减少次数比降低单次开销更有效。**

### 3. Benchmark 设计决定结论 — 缓存 shape 引用会制造虚假胜出

第一版 benchmark 中 pywin32 的 `move_shape` 赢了 VBA (18ms vs 28ms)，因为 shape 引用被提前缓存了。公平计入 find 后 pywin32 变成 65ms，VBA 仍是 22ms。**Benchmark 中隐含的状态缓存会严重扭曲结论。** 所有写操作必须包含完整的 find + modify 流程才有参考价值。

### 4. 没有全场景最优的 backend，只有最优的组合

- VBA 全量 inspect 快 17x，但单页 inspect 输给 C#
- C# 简单写操作快 3-5x，但全量读输给 VBA
- 根因：VBA 赢在"零 IPC 遍历"，C# 赢在"管道通信快"，两者优势来自不同维度，不可互相替代

### 5. "进程内 + 管道通信"才是理论终极方案

| 方案 | shape 访问 | 通信方式 | 综合 |
|------|-----------|---------|------|
| VBA (当前) | ✅ 进程内 0ms | ❌ App.Run 6ms | 🥇 实际最优 |
| C# COM Add-in (假设) | ✅ 进程内 0ms | ✅ 管道 0.3ms | 🏆 理论最优 |
| C# EXE (当前) | ❌ 跨进程 COM | ✅ 管道 0.3ms | 🥈 |
| pywin32 (当前) | ❌ 跨进程 COM | ❌ COM IPC 10ms | 🥉 |

但 C# Add-in 的部署复杂度（注册表注册、证书签名、VSTO 安装）远高于 VBA 的 `.bas` 文件导入，在离线/气隙场景中不现实。**工程上的最优解往往不是技术上的最优解。**

### 6. 实际影响：从"等一等"到"接近即时"

真实用户交互（inspect → 3 次修改 → 保存）的体感延时：

- pywin32: **3.07s** — 明显等待，打断思路
- C#: **582ms** — 略有感知，可接受
- VBA: **489ms** — 接近即时，无感知延迟

对于交互式编辑场景，**500ms 是用户体验的关键阈值**，VBA 和 C# 都在阈值内，pywin32 超出 6 倍。
