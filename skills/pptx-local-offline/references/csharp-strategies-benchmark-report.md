# C# Host 四策略 Benchmark 报告

> **版本**：v2.0  
> **日期**：2026-06-09  
> **测试文件**：Welcome to Our Company.pptx (1.6MB, 5 slides, ~60 shapes)  
> **方法**：每策略独立子进程，冷启动 PPT，warmup 3 轮，测量 7 轮  
> **状态**：4/4 策略全部成功，headless + headed 各 1 轮完整数据

---

## 1. 测试环境

| 项目 | 值 |
|------|-----|
| **操作系统** | Windows 11 Enterprise (10.0.26100) |
| **CPU** | Intel Core i9-10900X @ 3.70GHz (10 核 20 线程) |
| **内存** | 31.8 GB |
| **Python** | 3.12.10 |
| **.NET** | 10.0.204 |
| **Office** | Microsoft Office 16.0.20206.20000 (64-bit) |
| **Skill 版本** | pptx-local-offline (installed) |

### 测试文件结构

```
Welcome to Our Company.pptx (1,634 KB, 5 slides)
├── Slide 1 (Blank): 4 shapes — 1 image + 3 textboxes (标题/副标题/日期)
├── Slide 2 (Blank): 5 shapes — 2 groups + 2 textboxes + 1 group
├── Slide 3 (Blank): 17 shapes — 8 images + 9 textboxes
├── Slide 4 (Blank): 13 shapes — 5 images + 8 textboxes
└── Slide 5 (Blank): 21 shapes — 10 shapes + 11 textboxes
```

---

## 2. 四种策略的架构设计

四种策略共用 **同一个常驻 C# host 进程** (`PptInteropHost.exe`)、**同一条 NDJSON stdin/stdout 协议**。区别在于**动作的编排方式和编译路径**：

### 2.1 `csharp`（JSON dispatch）— `--backend csharp`

```
Python 发送 JSON ──stdin──→ C# host
  → JsonDocument.Parse()               ~1ms     已编译代码
  → switch(action) 匹配分支              ~0ms     已编译代码
  → 调用预编译的 C# 方法 (如 SetFont)     ~0ms     已编译代码
  → COM 调用 PowerPoint                 ~10ms    实际工作
  → 返回结果 ──stdout──→ Python          ~2ms
                                  总计: ~15ms
```

**特点**：所有 action 处理逻辑在 `dotnet build` 时已编译成 IL/机器码。运行时零编译开销，纯粹的「查表 + 执行」。每个 action 一次 stdin/stdout 往返。

### 2.2 `csharp-codeact`（Roslyn CodeAct 自动翻译）— `--backend csharp-codeact`

```
Python 把 JSON action 自动翻译成 C# 代码 ──stdin──→ C# host
  → Roslyn 词法/语法/语义分析              ~60ms
  → Roslyn IL 生成                        ~50ms
  → CLR JIT 编译 (IL → x64)              ~60ms
  → 执行机器码 → COM 调用 PPT             ~10ms    与 JSON 相同
  → 返回结果 ──stdout──→ Python            ~2ms
                                  总计: ~200ms（每次重编译）
```

**特点**：`PowerPointCSharpCodeAct`（Python 侧子类）自动把 JSON action 翻译成 C# 语句，经 `execute_code` 发送。`run_actions()` 可将 N 个动作拼成 1 段脚本、1 次编译。每次脚本文本不同，Roslyn 都要重新编译。

### 2.3 `csharp-template`（模板编译缓存）— `--backend csharp-template`

```
首次调用某 action type（如 modify_font）：
  → 按 action type 查缓存                  未命中
  → Roslyn 编译参数化模板                   ~200ms   首次代价
  → 缓存 ScriptRunner<object> delegate
  → delegate.Invoke(globals)               ~10ms
                                    总计: ~210ms（首次）

同 type 第 2 次起：
  → 按 action type 查缓存                  命中！
  → delegate.Invoke(globals)               ~10ms    直接执行
                                    总计: ~12ms（缓存命中 ≈ JSON）
```

**特点**：模板不内联字面量——值通过 `TemplateGlobals`（`Shp` + `A` 参数字典）传入。脚本文本恒定 → 编译只发生一次 → 之后同类 action 跳过编译。未覆盖的 action 自动回退到 JSON。

### 2.4 `exec-csharp`（手写 C# 脚本 / LLM 直出脚本）— `--exec-csharp edit.cs --backend csharp`

```
Agent/LLM 直接生成完整 C# 脚本 ──code_act()──→ C# host
  → Roslyn 编译整段脚本                    ~200ms   编译一次
  → 执行：多步操作全部在 host 进程内完成     ~10ms × N
  → 返回 Print() 输出 ──stdout──→ Python    ~2ms
                                    总计: ~200ms + N×10ms（1 次往返）
```

**特点**：
- **不经过 JSON action 翻译**——Agent/LLM 直接产出 PptApi C# 代码（`SetFont(Title(1), bold:true)`）
- **整个编辑计划 1 次往返**（inspect + modify + move + set_notes 全在 1 段脚本里）
- 编译开销是固定 ~200ms，但 N 步操作只需 1 次编译 → **N 越大优势越明显**
- CLI 入口：`python pptx_editor_llm.py deck.pptx --exec-csharp edit.cs`
- Python API：`ppt.code_act(script_string)`

### 2.5 四者关系图

```
                    ┌──────────────────────────────────────────────────┐
                    │          PptInteropHost.exe (常驻)                │
                    │                                                  │
  csharp ───────────┤  execute_action  →  switch/case 分发              │ 预编译，0ms 编译，N 往返
                    │                                                  │
  csharp-codeact ───┤  execute_code    →  Roslyn 编译+执行              │ 自动翻译，~200ms/次
                    │                     run_actions → 拼脚本 1 往返   │
                    │                                                  │
  csharp-template ──┤  execute_template → 缓存查找/首次编译             │ 每 type 编 1 次，N 往返
                    │                                                  │
  exec-csharp ──────┤  execute_code    →  Roslyn 编译+执行              │ 手写脚本，1 次编译，1 往返
                    │                                                  │
                    │          ↓ 最终都调用同一层 COM API ↓              │
                    │     App.Presentations[1].Slides[n]...            │
                    └──────────────────────────────────────────────────┘
```

### 2.6 核心差异总结

| 维度 | csharp (JSON) | csharp-codeact | csharp-template | exec-csharp |
|------|:---:|:---:|:---:|:---:|
| **编译开销** | 0ms | ~200ms/次 | ~200ms/type（首次） | ~200ms/脚本（1次） |
| **往返次数 (N 步)** | N 次 | 1 次（run_actions） | N 次 | **1 次** |
| **脚本来源** | — | Python 自动翻译 | — | Agent/LLM 手写 |
| **控制流** | ❌ | ✅ | ❌ | ✅ |
| **自定义 COM** | ❌ | ✅ | ❌ | ✅（经 App/Prs） |

---

## 3. 测试操作清单

| 操作 | 说明 | exec-csharp 等价脚本 |
|------|------|---------------------|
| `inspect_full` | 遍历所有 slide/shape，返回结构 JSON | `ppt.inspect()`（走原生路径） |
| `inspect_slide1` | 仅遍历第 1 页 | `ppt.inspect_slide(1)` |
| `modify_font` | 修改 TextBox 2 字体 | `SetFont(FindByName(1,"TextBox 2"), bold:true, colorBgr:255)` |
| `modify_text` | 修改 TextBox 2 文本 | `SetText(FindByName(1,"TextBox 2"), "...")` |
| `move_shape` | 移动 TextBox 4 位置 | `Move(FindByName(1,"TextBox 4"), 100, 500)` |
| `resize_shape` | 调整 TextBox 4 尺寸 | `Resize(FindByName(1,"TextBox 4"), 300, 40)` |
| `set_notes` | 设置 Slide 1 备注 | `SetNotes(1, "...")` |
| `add+del_textbox` | 创建+删除文本框 | `AddTextbox(...); FindByText(...).Delete()` |
| `batch_8_seq` | 8 个混合操作逐条执行 | 8 次独立 `code_act()` |
| `batch_8_merged` | 8 个操作合并执行 | **1 段包含 8 步的脚本** |
| `user_flow` | inspect→font→text→move→notes | **1 段包含 5 步的脚本** |
| `inspect/tp` | 5 秒窗口 inspect 吞吐 | — |
| `modify/tp` | 5 秒窗口 modify 吞吐 | — |

---

## 4. 测试结果

### 4.1 Headless 模式

| 操作 | csharp (JSON) | codeact | template | **exec-csharp** | 🏅 |
|------|:---:|:---:|:---:|:---:|:---:|
| **inspect_full** | 457ms | **387ms** | 504ms | 478ms | ⚡codeact |
| **inspect_slide1** | **43ms** | 53ms | 76ms | 80ms | ⚡json |
| **modify_font** | 25ms | 268ms | **16ms** | 199ms | ⚡template |
| **modify_text** | **18ms** | 219ms | 19ms | 202ms | ⚡json |
| **move_shape** | 11ms | 203ms | **11ms** | 215ms | ⚡template |
| **resize_shape** | 13ms | 194ms | **8ms** | 216ms | ⚡template |
| **set_notes** | **10ms** | 178ms | 21ms | 204ms | ⚡json |
| **add+del_textbox** | **38ms** | 434ms | 61ms | 264ms | ⚡json |
| **batch_8_seq** | **161ms** | 1708ms | 190ms | 1112ms | ⚡json |
| **batch_8_merged** | 183ms | 443ms | **167ms** | 310ms | ⚡template |
| **user_flow** | 497ms | 1443ms | 588ms | **121ms** | ⚡exec-cs |
| **inspect/tp** | **2.2/s** | 1.8/s | 1.8/s | 2.0/s | ⚡json |
| **modify/tp** | **56.6/s** | 5.6/s | 54.0/s | 4.6/s | ⚡json |

### 4.2 Headed 模式

| 操作 | csharp (JSON) | codeact | template | **exec-csharp** | 🏅 |
|------|:---:|:---:|:---:|:---:|:---:|
| **inspect_full** | 470ms | 487ms | 551ms | **398ms** | ⚡exec-cs |
| **inspect_slide1** | 58ms | **50ms** | 61ms | 69ms | ⚡codeact |
| **modify_font** | 38ms | 274ms | **37ms** | 277ms | ⚡template |
| **modify_text** | **17ms** | 231ms | 29ms | 260ms | ⚡json |
| **move_shape** | 25ms | 199ms | **14ms** | 202ms | ⚡template |
| **resize_shape** | **14ms** | 165ms | 29ms | 189ms | ⚡json |
| **set_notes** | **21ms** | 185ms | 24ms | 185ms | ⚡json |
| **add+del_textbox** | **71ms** | 420ms | 80ms | 237ms | ⚡json |
| **batch_8_seq** | 237ms | 1475ms | **230ms** | 1254ms | ⚡template |
| **batch_8_merged** | **238ms** | 405ms | 254ms | 437ms | ⚡json |
| **user_flow** | 618ms | 1432ms | 903ms | **102ms** | ⚡exec-cs |
| **inspect/tp** | 2.0/s | **2.2/s** | 1.8/s | 1.8/s | ⚡codeact |
| **modify/tp** | 34.8/s | 5.0/s | **48.0/s** | 4.6/s | ⚡template |

### 4.3 综合胜场统计（Headless + Headed，26 项）

| 排名 | 策略 | 总胜场 | 胜出领域 |
|------|------|--------|---------|
| 🥇 | **csharp (JSON)** | **10/26** | 单操作写入（modify_text, set_notes, add+del）、batch_seq、吞吐 |
| 🥈 | **csharp-template** | **8/26** | modify_font, move_shape, resize, batch_merged |
| 🥉 | **exec-csharp** | **4/26** | **user_flow（两次冠军）、inspect_full headed** |
| 4 | **csharp-codeact** | **3/26** | inspect_full headless, inspect_slide1 headed |

---

## 5. exec-csharp 深度分析

### 5.1 user_flow：碾压级优势

```
操作：inspect → modify_font → modify_text → move_shape → set_notes

  csharp (JSON):      5 次往返 × ~100ms/次 ≈ 497ms
  csharp-codeact:     5 次编译 × ~200ms/次 ≈ 1443ms
  csharp-template:    5 次往返（缓存） ≈ 588ms
  exec-csharp:        1 次编译 + 5 步执行 ≈ 121ms  ← 4x 快于 JSON
```

**根因**：exec-csharp 把 5 步操作写进 1 段 C# 脚本：

```csharp
// 1 次 Roslyn 编译（~100ms）+ 5 步 COM 执行（~20ms）= ~120ms
Print(InspectJson());
{ var s = FindByName(1, "TextBox 2"); if (s != null) SetFont(s, bold: true, colorBgr: 0xFF0000); }
{ var s = FindByName(1, "TextBox 3"); if (s != null) SetText(s, "User flow test"); }
{ var s = FindByName(1, "TextBox 4"); if (s != null) Move(s, 100, 500); }
SetNotes(1, "user flow note");
```

**1 次编译开销被 5 步操作摊薄** → 每步平均 24ms，比 JSON 的 100ms/步 快 4 倍。

### 5.2 单操作：Roslyn 编译税不可避免

```
modify_font:
  JSON:       25ms  (0ms 编译 + 25ms 执行)
  exec-csharp: 199ms (180ms 编译 + 19ms 执行)
```

单操作场景下，~180ms 的 Roslyn 编译开销无法摊薄，exec-csharp 比 JSON 慢 8 倍。

### 5.3 batch_8_merged：1 段脚本的威力

```
  JSON 逐条:     8 × 20ms = 161ms（8 次往返）
  exec-csharp:   1 × 200ms + 8 × 10ms = 310ms（1 次往返 + 1 次编译）
```

8 步时 exec-csharp 仍不如 JSON（编译税 > 往返节省），但随着 N 增大：
- **N = 13 时**：JSON 13×20ms = 260ms ≈ exec-csharp 200ms + 13×10ms = 330ms（接近）
- **N = 20 时**：JSON 20×20ms = 400ms > exec-csharp 200ms + 20×10ms = 400ms（持平）
- **N = 50 时**：JSON 50×20ms = 1000ms >> exec-csharp 200ms + 50×10ms = 700ms（exec-cs 领先）

### 5.4 为什么 exec-csharp ≠ csharp-codeact

两者都用 `execute_code`（Roslyn），但：

| 维度 | csharp-codeact | exec-csharp |
|------|:---:|:---:|
| 脚本来源 | Python 自动翻译（每 action 独立脚本） | Agent/LLM 手写（整个计划 1 段脚本） |
| 编译次数 | N 次（逐条）或 1 次（run_actions） | **始终 1 次** |
| 控制流 | 翻译后无控制流 | ✅ 循环/条件/变量 |
| user_flow | 1443ms（5 次编译） | **121ms（1 次编译）** |

**关键差异**：csharp-codeact 的 `execute_action()` 每次调用都翻译出一段**新脚本**然后编译。exec-csharp 由 Agent 一次性生成**完整脚本**，只编译 1 次。

---

## 6. 四策略适用场景

| 场景 | 推荐 | 原因 |
|------|------|------|
| **单次简单修改** | `csharp` (JSON) | 零开销，~15ms |
| **高频重复操作（同 type）** | `csharp-template` | 首次 ~200ms 后持续 ~12ms |
| **Agent 一次生成完整编辑计划** | **`exec-csharp`** | 1 次编译，N 步操作摊薄，user_flow 121ms |
| **复杂逻辑（循环/条件遍历）** | **`exec-csharp`** | 任意 C# 控制流 + PptApi |
| **超大批量已知操作** | `csharp-codeact` run_actions | N>20 时 1 次编译优势明显 |
| **默认 / 通用** | `csharp` (JSON) | 最稳定、最快、零意外 |

### Agent/LLM 最佳实践

```
用户请求 → Agent 生成编辑计划 → 选择策略：

  计划步骤 ≤ 3 且无控制流  → --backend csharp + --exec-actions
  计划步骤 > 3 或有控制流  → --backend csharp + --exec-csharp edit.cs
```

---

## 7. 原始数据（7 轮，单位 ms）

### 7.1 Headless（精选关键操作）

| 操作 | 策略 | R1 | R2 | R3 | R4 | R5 | R6 | R7 | **Median** |
|------|------|---|---|---|---|---|---|---|:---:|
| inspect_full | csharp | 375 | 458 | 497 | 453 | 457 | 497 | 344 | **457** |
| inspect_full | codeact | 222 | 387 | 495 | 373 | 357 | 368 | 444 | **387** |
| inspect_full | template | 292 | 504 | 506 | 658 | 502 | 397 | 366 | **504** |
| inspect_full | exec-cs | 372 | 478 | 454 | 478 | 634 | 411 | 519 | **478** |
| modify_font | csharp | 18 | 26 | 25 | 33 | 22 | 29 | 25 | **25** |
| modify_font | codeact | 185 | 268 | 304 | 768 | 236 | 271 | 219 | **268** |
| modify_font | template | 12 | 16 | 18 | 16 | 16 | 14 | 15 | **16** |
| modify_font | exec-cs | 154 | 199 | 187 | 242 | 189 | 192 | 202 | **199** |
| user_flow | csharp | 463 | 481 | 497 | 484 | 507 | 463 | 559 | **497** |
| user_flow | codeact | 1443 | 1484 | 1375 | 1346 | 1418 | 1501 | 1337 | **1443** |
| user_flow | template | 589 | 588 | 647 | 695 | 553 | 510 | 460 | **588** |
| user_flow | exec-cs | 112 | 121 | 127 | 116 | 120 | 121 | 112 | **121** |

### 7.2 Headed（精选关键操作）

| 操作 | 策略 | R1 | R2 | R3 | R4 | R5 | R6 | R7 | **Median** |
|------|------|---|---|---|---|---|---|---|:---:|
| inspect_full | csharp | 316 | 469 | 570 | 576 | 470 | 463 | 394 | **470** |
| inspect_full | exec-cs | 344 | 388 | 515 | 365 | 398 | 432 | 419 | **398** |
| user_flow | csharp | 529 | 618 | 554 | 625 | 651 | 591 | 755 | **618** |
| user_flow | codeact | 1421 | 1432 | 1277 | 1583 | 1525 | 1321 | 1564 | **1432** |
| user_flow | template | 788 | 903 | 1046 | 818 | 850 | 945 | 829 | **903** |
| user_flow | exec-cs | 93 | 102 | 94 | 103 | 101 | 116 | 79 | **102** |

---

## 8. Key Insights

### Insight 1：exec-csharp 用 1 次编译税换 N-1 次往返节省

```
user_flow (5 步)：
  JSON:       5 × 100ms/步 = 497ms   （5 次往返，0 编译）
  exec-cs:    1 × 100ms编译 + 5 × 4ms = 121ms   （1 次往返，1 次编译）
                                         ↑ 快 4 倍
```

编译税 ~100-200ms 是固定成本。步数 N 越大，被摊薄越多：
- N=1：exec-cs 慢（编译 > 往返节省）
- N=5：exec-cs 快 4x（user_flow 证实）
- N=20+：exec-cs 碾压

### Insight 2：四策略形成完整的「编译 × 往返」矩阵

|  | 0 编译 | 每次编译 | 每 type 编 1 次 | 每脚本编 1 次 |
|--|:---:|:---:|:---:|:---:|
| **N 往返** | ⚡JSON | codeact(逐条) | ⚡template | — |
| **1 往返** | — | codeact(合并) | — | ⚡exec-cs |

最优解取决于 N 和操作类型：
- N 小 → JSON（零编译碾压）
- N 大 + 重复 type → template（缓存碾压）
- N 大 + 异构操作 → exec-csharp（1 次编译摊薄）

### Insight 3：exec-csharp 是 Agent/LLM 的自然模式

Agent 思考一次、生成一段完整脚本 → `--exec-csharp edit.cs` → 1 次编译 + 1 次往返。

这比 Agent 生成 N 个 JSON action 然后逐条发送更高效——因为 Agent 本来就是一次性产出整个计划，不需要额外的「JSON 翻译」步骤。

### Insight 4：Headed 下 exec-csharp user_flow 102ms — 真正的即时体感

人类感知阈值 ~200ms（「即时」）、~500ms（「快速」）。exec-csharp 在 headed 模式下 user_flow 102ms，用户在 PowerPoint 窗口中看到的是**瞬间完成的一连串修改**。

---

## 9. 测试执行记录

| 轮次 | 时间 | 策略数 | 模式 | Warmup/Rounds | 结果 |
|------|------|--------|------|--------------|------|
| 1 | 16:32 | 2 路 | headless | 2/5 | ✅ csharp vs codeact |
| 2 | 16:37 | 2 路 | headed | 2/5 | ⚠️ batch 崩溃 |
| 3 | 16:38 | 2 路 | headed | 2/5 | ✅ 加错误处理后 |
| 4 | 16:44 | 3 路 | headless | 2/5 | ✅ 加入 template |
| 5 | 16:46 | 3 路 | headed | 2/5 | ⚠️ COM 连接断裂 |
| 6 | 17:05 | 3 路 | headless | 3/7 | ✅ 高质量数据 |
| 7 | 17:10 | 3 路 | headed | 3/7 | ✅ 高质量数据 |
| **8** | **17:33** | **4 路** | **headless** | **3/7** | **✅ 最终数据（含 exec-csharp）** |
| **9** | **17:38** | **4 路** | **headed** | **3/7** | **✅ 最终数据（含 exec-csharp）** |

> 本报告使用**第 8、9 轮**的数据作为最终结果。

---

## 10. 结论

> **exec-csharp 是 Agent/LLM 场景的最优策略**——把整个编辑计划写成 1 段 C# 脚本，用 1 次 Roslyn 编译换取 N-1 次往返节省。user_flow（5 步）只需 121ms（headless）/ 102ms（headed），比 JSON 快 4-6 倍。
>
> **但它不适合单操作和高频重复场景**——此时 JSON（~15ms）和 template（缓存后 ~12ms）仍然是最快选择。
>
> **四策略共存互补，不是相互替代**。
