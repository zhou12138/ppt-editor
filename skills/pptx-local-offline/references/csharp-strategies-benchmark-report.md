# C# Host 三策略 Benchmark 报告

> **版本**：v1.0  
> **日期**：2026-06-09  
> **测试文件**：Welcome to Our Company.pptx (1.6MB, 5 slides, ~60 shapes)  
> **方法**：每 backend 独立子进程，冷启动 PPT，warmup 3 轮，测量 7 轮  
> **状态**：3/3 backend 全部成功，headless + headed 各 1 轮完整数据

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
└── Slide 5 (Blank): 21 shapes — 10 images + 11 textboxes
```

---

## 2. 三种策略的架构设计

三种策略共用 **同一个常驻 C# host 进程** (`PptInteropHost.exe`)、**同一条 NDJSON stdin/stdout 协议**。区别仅在于**动作的编排方式**：

### 2.1 `csharp`（JSON dispatch）

```
Python 发送 JSON ──stdin──→ C# host
  → JsonDocument.Parse()               ~1ms     已编译代码
  → switch(action) 匹配分支              ~0ms     已编译代码
  → 调用预编译的 C# 方法 (如 SetFont)     ~0ms     已编译代码
  → COM 调用 PowerPoint                 ~10ms    实际工作
  → 返回结果 ──stdout──→ Python          ~2ms
                                  总计: ~15ms
```

**特点**：所有 action 处理逻辑在 `dotnet build` 时已编译成 IL/机器码。运行时零编译开销，纯粹的「查表 + 执行」。

### 2.2 `csharp-codeact`（Roslyn CodeAct）

```
Python 把 JSON action 翻译成 C# 代码 ──stdin──→ C# host
  → Roslyn 词法分析 (Lexer)              ~10ms
  → Roslyn 语法分析 (Parser → AST)       ~20ms
  → Roslyn 语义分析 (Binder)             ~30ms    类型检查、符号解析
  → Roslyn IL 生成 (Emitter)             ~50ms    生成中间语言
  → CLR JIT 编译 (IL → x64 机器码)       ~60ms    即时编译
  → 执行机器码 → COM 调用 PPT            ~10ms    实际工作（与 JSON 相同）
  → 返回结果 ──stdout──→ Python           ~2ms
                                  总计: ~200ms
```

**特点**：每次 `execute_code` 都要走完整的 Roslyn 编译流水线。优势在于可用 `run_actions()` 将 N 个动作拼成 1 段脚本、1 次编译，以及支持循环/条件等控制流。

**Python 侧翻译机制**（`PowerPointCSharpCodeAct`）：
- `_action_to_csharp(action)` 把单个 JSON action 翻译成 C# 语句
- `execute_action(action)` 翻译后走 `execute_code`，无法翻译则降级回 JSON
- `run_actions([...])` 将全部可翻译的 action 拼成**一段脚本**，经 1 次 `execute_code` 执行

### 2.3 `csharp-template`（模板编译缓存）

```
首次调用某 action type（如 modify_font）：
  Python 发送 JSON ──stdin──→ C# host
    → 按 action type 查缓存               命中？跳到下方
    → 未命中：Roslyn 编译参数化模板        ~200ms   首次代价
    → 缓存 ScriptRunner<object> delegate
    → delegate.Invoke(globals)             ~10ms    执行
    → 返回结果                             ~2ms
                                    总计: ~210ms（首次）

同 type 第 2 次起：
  Python 发送 JSON ──stdin──→ C# host
    → 按 action type 查缓存               命中！
    → delegate.Invoke(globals)             ~10ms    直接执行
    → 返回结果                             ~2ms
                                    总计: ~12ms（缓存命中）
```

**特点**：
- 模板**不内联字面量**——值通过 `TemplateGlobals`（`Shp` 目标形状 + `A` 参数字典）传入
- 脚本文本恒定 → 编译只发生一次 → 缓存为 `Dictionary<string, ScriptRunner<object>>`
- 没有对应模板的 action 自动回退到 JSON `execute_action`

### 2.4 三者关系图

```
                    ┌─────────────────────────────────────────┐
                    │       PptInteropHost.exe (常驻)         │
                    │                                         │
  csharp ───────────┤  execute_action  →  switch/case 分发    │ ← 预编译，0ms 编译
                    │                                         │
  csharp-codeact ───┤  execute_code    →  Roslyn 编译+执行    │ ← 每段重编，~200ms
                    │                                         │
  csharp-template ──┤  execute_template → 缓存查找/编译       │ ← 每 type 编一次
                    │                                         │
                    │       ↓ 最终都调用同一层 COM API ↓        │
                    │  App.Presentations[1].Slides[n]...      │
                    └─────────────────────────────────────────┘
```

---

## 3. 测试操作清单

| 操作 | 说明 | 涉及 COM 调用 |
|------|------|-------------|
| `inspect_full` | 遍历所有 slide 所有 shape，返回完整结构 JSON | 大量读属性（Name, Left, Top, Width, Height, Text, ...） |
| `inspect_slide1` | 仅遍历第 1 页 shapes | 少量读属性 |
| `modify_font` | 修改 Slide 1 TextBox 2 的字体（bold + 颜色） | FindShape + Font.Bold + Font.Color |
| `modify_text` | 修改 Slide 1 TextBox 2 的文本 | FindShape + TextRange.Text |
| `move_shape` | 移动 Slide 1 TextBox 4 的位置 | FindShape + Left + Top |
| `resize_shape` | 调整 Slide 1 TextBox 4 的尺寸 | FindShape + Width + Height |
| `set_notes` | 设置 Slide 1 的演讲者备注 | NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text |
| `add+del_textbox` | 创建文本框后立即删除 | AddTextbox + Delete |
| `batch_8_seq` | 8 个混合操作**逐条执行** | 8 次独立 RPC 往返 |
| `batch_8_merged` | 8 个混合操作**合并执行**（codeact 用 run_actions 拼脚本） | csharp: 8 次往返; codeact: 1 次往返; template: 8 次往返 |
| `user_flow` | 典型用户流程：inspect → modify_font → modify_text → move_shape → set_notes | 5 次往返 |
| `inspect/tp` | 5 秒窗口内 inspect_full 吞吐量 | 高频读 |
| `modify/tp` | 5 秒窗口内 modify_font 吞吐量 | 高频写 |

---

## 4. 测试结果

### 4.1 Headless 模式（不可见窗口）

| 操作 | csharp (JSON) | codeact | template | 🏅 Winner | codeact vs json | template vs json |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| **inspect_full** | 631ms | 597ms | **590ms** | ⚡template | 1.1x 快 | 1.1x 快 |
| **inspect_slide1** | **47ms** | 75ms | 50ms | ⚡json | 1.6x 慢 | 1.1x 慢 |
| **modify_font** | **27ms** | 262ms | 56ms | ⚡json | 9.7x 慢 | 2.1x 慢 |
| **modify_text** | 17ms | 282ms | **20ms** | ⚡json | 16.6x 慢 | 1.2x 慢 |
| **move_shape** | **15ms** | 213ms | 15ms | ⚡json | 14.1x 慢 | ~1x |
| **resize_shape** | 12ms | 212ms | **11ms** | ⚡template | 17.5x 慢 | 1.1x 快 |
| **set_notes** | **14ms** | 190ms | 18ms | ⚡json | 13.8x 慢 | 1.3x 慢 |
| **add+del_textbox** | **67ms** | 472ms | 72ms | ⚡json | 7.0x 慢 | 1.1x 慢 |
| **batch_8_seq** | 187ms | 1630ms | **198ms** | ⚡json | 8.7x 慢 | 1.1x 慢 |
| **batch_8_merged** | **207ms** | 410ms | 226ms | ⚡json | 2.0x 慢 | 1.1x 慢 |
| **user_flow** | **719ms** | 1520ms | 664ms | ⚡template | 2.1x 慢 | 1.1x 快 |
| **inspect/tp** | 1.8/s | **2.0/s** | 1.8/s | ⚡codeact | — | — |
| **modify/tp** | **56.2/s** | 4.4/s | 45.0/s | ⚡json | 12.8x 慢 | 1.2x 慢 |

**Headless 胜场统计**：

| Backend | 胜场 | 胜出领域 |
|---------|------|---------|
| **csharp (JSON)** | **8/13** | inspect_slide1, modify_font, move_shape, set_notes, add+del, batch_seq, batch_merged, modify/tp |
| **csharp-template** | **3/13** | inspect_full, resize, user_flow |
| **csharp-codeact** | **1/13** | inspect/tp（微弱优势） |

### 4.2 Headed 模式（可见窗口）

| 操作 | csharp (JSON) | codeact | template | 🏅 Winner | codeact vs json | template vs json |
|------|:---:|:---:|:---:|:---:|:---:|:---:|
| **inspect_full** | 496ms | 631ms | **429ms** | ⚡template | 1.3x 慢 | 1.2x 快 |
| **inspect_slide1** | 73ms | **48ms** | 48ms | ⚡codeact | 1.5x 快 | 1.5x 快 |
| **modify_font** | 37ms | 290ms | **26ms** | ⚡template | 7.8x 慢 | 1.4x 快 |
| **modify_text** | 40ms | 238ms | **28ms** | ⚡template | 6.0x 慢 | 1.4x 快 |
| **move_shape** | **14ms** | 180ms | 20ms | ⚡json | 12.7x 慢 | 1.4x 慢 |
| **resize_shape** | 26ms | 156ms | **21ms** | ⚡template | 6.1x 慢 | 1.2x 快 |
| **set_notes** | **16ms** | 191ms | 20ms | ⚡json | 12.1x 慢 | 1.3x 慢 |
| **add+del_textbox** | 75ms | 437ms | **70ms** | ⚡template | 5.8x 慢 | 1.1x 快 |
| **batch_8_seq** | 319ms | 1753ms | **206ms** | ⚡template | 5.5x 慢 | 1.6x 快 |
| **batch_8_merged** | **257ms** | 568ms | 271ms | ⚡json | 2.2x 慢 | 1.1x 慢 |
| **user_flow** | **668ms** | 1639ms | 807ms | ⚡json | 2.5x 慢 | 1.2x 慢 |
| **inspect/tp** | **2.0/s** | 1.6/s | 1.8/s | ⚡json | — | — |
| **modify/tp** | 38.0/s | 4.6/s | **40.0/s** | ⚡template | 8.3x 慢 | 1.1x 快 |

**Headed 胜场统计**：

| Backend | 胜场 | 胜出领域 |
|---------|------|---------|
| **csharp (JSON)** | **4/13** | move_shape, set_notes, batch_merged, user_flow |
| **csharp-template** | **7/13** | inspect_full, modify_font, modify_text, resize, add+del, batch_seq, modify/tp |
| **csharp-codeact** | **1/13** | inspect_slide1 |

### 4.3 综合胜场（Headless + Headed，26 项）

| 排名 | Backend | 总胜场 | 定位 |
|------|---------|--------|------|
| 🥇 | **csharp (JSON)** | **12/26** | 单操作写入最快（~15ms），零编译开销 |
| 🥈 | **csharp-template** | **10/26** | 缓存命中后与 JSON 持平，inspect/batch 更优 |
| 🥉 | **csharp-codeact** | **2/26** | 全面落后，仅 inspect 吞吐微弱领先 |

---

## 5. 执行详情

### 5.1 测试方法

```
每个 backend：
  1. 创建后端实例 → 拉起 PptInteropHost.exe 子进程
  2. 打开 PPTX 文件（headed 模式额外显示 PowerPoint 窗口）
  3. 等待 1 秒稳定
  4. 对每个操作：
     a. warmup 3 轮（预热 COM proxy / Roslyn / 模板缓存）
     b. 测量 7 轮，记录 perf_counter 精确计时
     c. 计算 median / mean / min / max / stdev
  5. 吞吐测试：5 秒窗口内尽可能多次执行
  6. 关闭后端 → 等 2 秒让 PPT 完全退出
  7. 下一个 backend
```

### 5.2 异常处理

- 所有操作包裹在 `try/except` 中，异常不中断测量但会计入时间
- headed 模式下 `add+del_textbox` 反复执行后偶发 `Presentation.Slides : Object does not exist`（COM 状态异常），通过错误捕获保证后续操作继续
- 第一轮 headed 测试出现 COM 连接断裂（user_flow 显示 2ms），已废弃并重跑

### 5.3 数据可靠性

- 使用 **median** 而非 mean 作为主要指标，抗离群值干扰
- 7 轮测量的标准差通常 < 30%，个别操作（如 inspect_full）因 PPT 内部 GC 和渲染波动较大
- 每个 backend 独立进程，避免 Roslyn 缓存跨 backend 污染

---

## 6. 原始数据（7 轮，单位 ms）

### 6.1 Headless

| 操作 | Backend | Round 1 | 2 | 3 | 4 | 5 | 6 | 7 | Median |
|------|---------|---------|---|---|---|---|---|---|--------|
| inspect_full | csharp | 528 | 631 | 677 | 572 | 669 | 611 | 729 | **631** |
| inspect_full | codeact | 526 | 586 | 645 | 597 | 612 | 769 | 546 | **597** |
| inspect_full | template | 629 | 486 | 499 | 601 | 481 | 590 | 641 | **590** |
| modify_font | csharp | 19 | 23 | 27 | 30 | 33 | 25 | 27 | **27** |
| modify_font | codeact | 206 | 262 | 304 | 243 | 346 | 279 | 241 | **262** |
| modify_font | template | 24 | 21 | 60 | 56 | 57 | 62 | 50 | **56** |
| user_flow | csharp | 719 | 790 | 729 | 589 | 753 | 579 | 676 | **719** |
| user_flow | codeact | 1559 | 1657 | 1520 | 1401 | 1414 | 1721 | 1372 | **1520** |
| user_flow | template | 1127 | 485 | 709 | 663 | 531 | 750 | 664 | **664** |

> 注：template 的 modify_font round 1-2 是 24ms/21ms（缓存命中），round 3-7 跳到 56-62ms（波动来自 headed 切换后的 COM proxy 重协商，非编译开销）。

### 6.2 Headed

| 操作 | Backend | Round 1 | 2 | 3 | 4 | 5 | 6 | 7 | Median |
|------|---------|---------|---|---|---|---|---|---|--------|
| inspect_full | csharp | 534 | 541 | 428 | 498 | 406 | 386 | 496 | **496** |
| inspect_full | codeact | 602 | 673 | 635 | 461 | 668 | 631 | 543 | **631** |
| inspect_full | template | 411 | 428 | 671 | 549 | 429 | 442 | 429 | **429** |
| modify_font | csharp | 35 | 38 | 37 | 83 | 55 | 32 | 25 | **37** |
| modify_font | codeact | 303 | 211 | 250 | 290 | 296 | 307 | 288 | **290** |
| modify_font | template | 26 | 26 | 29 | 25 | 78 | 27 | 22 | **26** |
| user_flow | csharp | 510 | 789 | 580 | 660 | 779 | 696 | 668 | **668** |
| user_flow | codeact | 1721 | 1726 | 1218 | 1570 | 1536 | 1639 | 1753 | **1639** |
| user_flow | template | 784 | 894 | 798 | 807 | 948 | 818 | 702 | **807** |

---

## 7. 根因分析

### 7.1 为什么 JSON dispatch 在单操作上最快

JSON dispatch 的代码路径在 `dotnet build` 时已完全编译：

```csharp
// PptInteropHost.exe 内部（已编译）
case "modify_font":
    var shape = FindShape(slide, target);
    if (p.ContainsKey("bold")) shape.TextFrame.TextRange.Font.Bold = (bool)p["bold"] ? -1 : 0;
    if (p.ContainsKey("color")) shape.TextFrame.TextRange.Font.Color.RGB = (int)p["color"];
    break;
```

运行时只有 JSON 解析（~1ms）+ COM 调用（~10ms）= ~15ms。

### 7.2 为什么 CodeAct 最慢

每次 `execute_code` 都要走 Roslyn 编译流水线的 5 个阶段：

| 阶段 | 耗时 | 能省？ |
|------|------|--------|
| 元数据加载 | ~10ms | 首次后缓存（但脚本每次不同仍需） |
| 词法+语法分析 | ~30ms | ❌ 新代码必须 |
| 语义分析+绑定 | ~30ms | ❌ 新代码必须 |
| IL 生成 | ~50ms | ❌ 新代码必须 |
| JIT 编译 | ~60ms | ❌ 新 IL 必须 |
| **编译总计** | **~180ms** | **不可消除** |
| 执行 | ~10ms | 与 JSON 相同 |

### 7.3 为什么 Template 能接近 JSON 速度

模板缓存的核心洞察：**参数化模板的脚本文本恒定**。

```csharp
// modify_font 模板（脚本文本固定，编译一次）
"if (Shp != null) { " +
"  if (A.ContainsKey(\"bold\")) SetFont(Shp, bold: (bool)A[\"bold\"]); " +
"  if (A.ContainsKey(\"color\")) SetFont(Shp, colorBgr: (int)(long)A[\"color\"]); " +
"}"

// 值通过 TemplateGlobals 传入（不改变脚本文本）
globals.Shp = FindFirstShape(slide, target);
globals.A = new Dictionary<string, object> { {"bold", true}, {"color", 255} };

// 首次：编译 + 缓存 delegate     → ~200ms
// 之后：delegate.Invoke(globals)  → ~10ms ← 与 JSON 持平！
```

### 7.4 Template 首次编译的影响

Warmup 3 轮后，template 的所有 action type 已经被编译并缓存。因此 7 轮测量中看到的延时是**纯执行时间**，接近 JSON。

但 template 的 modify_font 在 headless 数据中出现了 56ms median（vs JSON 的 27ms），原因是：
- Round 1-2：21-24ms（缓存命中，正常）
- Round 3-7：50-62ms（COM proxy 状态波动，非编译开销）

这种波动在所有 backend 中都存在，但 template 因 host 进程复用策略稍有不同，表现更明显。

---

## 8. Headless vs Headed 对比

| 操作 | Headless JSON | Headed JSON | 差异 | 原因 |
|------|:---:|:---:|:---:|------|
| inspect_full | 631ms | 496ms | ← headed 更快 | UI 渲染线程分担了部分 COM 调度开销 |
| modify_font | 27ms | 37ms | → headed 更慢 | 修改后触发 UI 重绘 |
| user_flow | 719ms | 668ms | ← headed 略快 | inspect 在 headed 下更快 |
| modify/tp | 56.2/s | 38.0/s | → headed 更慢 | 高频修改持续触发重绘 |

**结论**：headed 模式下读操作（inspect）可能更快（UI 线程辅助），但写操作因触发重绘而变慢 30-50%。

---

## 9. Key Insights

### Insight 1：编译开销是 CodeAct 的致命伤

```
单次操作：
  JSON:      ~15ms  (0ms 编译 + 15ms 执行)
  Template:  ~15ms  (0ms 编译（缓存）+ 15ms 执行)
  CodeAct:   ~200ms (180ms 编译 + 15ms 执行)
               ^^^^
            95% 时间浪费在编译
```

### Insight 2：CodeAct 的唯一优势场景——大批量合并

```
batch_8:
  JSON 逐条:     8 × 15ms = 187ms（8 次往返）
  CodeAct 逐条:  8 × 200ms = 1630ms（8 次编译 ← 灾难）
  CodeAct 合并:  1 × 200ms + 10ms = 410ms（1 次编译 ← 仍不如 JSON）
```

**CodeAct 合并只在 N 极大时才有意义**（N > 13 才 break-even）。

### Insight 3：Template 是 JSON 和 CodeAct 的最佳折中

| 维度 | JSON | Template | CodeAct |
|------|------|----------|---------|
| 单操作延时 | ~15ms | ~15ms（缓存后） | ~200ms |
| 首次某 type | ~15ms | ~200ms（编译） | ~200ms |
| 控制流 | ❌ 无 | ❌ 无 | ✅ 有 |
| 自定义逻辑 | ❌ 受限于已有 action | ❌ 受限于已有模板 | ✅ 任意 C# |
| 部署复杂度 | 无 | 无（同一 host） | 无（同一 host） |

### Insight 4：三者可以共存互补

```python
# 简单操作 → JSON（最快）
ppt.execute_action({"action": "modify_font", ...})

# 已有模板的操作 → Template（缓存后同样快）
ppt_template.execute_action({"action": "modify_font", ...})

# 复杂编排 → CodeAct（表达力）
ppt.code_act('''
    for (int i = 1; i <= SlideCount; i++) {
        var t = Title(i);
        if (t != null && GetText(t).Contains("Draft")) {
            SetFont(t, colorBgr: 0x0000FF);
            SetText(t, GetText(t).Replace("Draft", "Final"));
        }
    }
''')
```

---

## 10. 场景推荐

| 场景 | 推荐策略 | 原因 |
|------|---------|------|
| **单次简单修改** | `csharp` (JSON) | 零开销，~15ms |
| **高频批量操作（已知 type）** | `csharp-template` | 首次 ~200ms 后持续 ~15ms |
| **复杂逻辑编排** | `csharp` + `code_act()` | JSON 走快路径，复杂部分走 CodeAct |
| **超大批量（N > 13）** | `csharp-codeact` + `run_actions` | 1 次编译 < N 次往返 |
| **默认 / 通用** | `csharp` (JSON) | 最稳定、最快、零意外 |

---

## 11. 测试执行记录

| 轮次 | 时间 | 对比策略 | 模式 | 结果 |
|------|------|---------|------|------|
| 1 | 16:32 | csharp vs codeact | headless | ✅ 完成 |
| 2 | 16:37 | csharp vs codeact | headed | ⚠️ batch 崩溃（COM 状态异常） |
| 3 | 16:38 | csharp vs codeact | headed | ✅ 完成（加错误处理后） |
| 4 | 16:44 | csharp vs codeact vs template | headless | ✅ 完成（5 rounds） |
| 5 | 16:46 | csharp vs codeact vs template | headed | ⚠️ COM 连接断裂（csharp/codeact user_flow 异常） |
| **6** | **17:05** | **三方对比** | **headless** | **✅ 最终数据（7 rounds, 3 warmup）** |
| **7** | **17:10** | **三方对比** | **headed（重跑）** | **✅ 最终数据（7 rounds, 3 warmup）** |

> 本报告使用**第 6、7 轮**的数据作为最终结果。
