"""
PPTX 编辑器 — LLM 意图解析 + COM 执行 (Windows Only)
支持三种执行模式: LLM自然语言、本地脚本执行、本地JSON动作执行
需要: pip install pywin32 requests + Microsoft Office + OpenAI-compatible API

用法:
  # LLM 模式 (需要 API key)
  python pptx_editor_llm.py <pptx文件> "把标题改成红色"
  python pptx_editor_llm.py <pptx文件> --interactive
  python pptx_editor_llm.py <pptx文件> "指令" --dry-run
  python pptx_editor_llm.py <pptx文件> "指令" --output out.pptx

  # 本地脚本执行 (无需 API)
  python pptx_editor_llm.py <pptx文件> --exec-script edit.py
  python pptx_editor_llm.py <pptx文件> --exec-script edit.py --output out.pptx

  # 本地JSON动作执行 (无需 API)
  python pptx_editor_llm.py <pptx文件> --exec-actions '[{"action":"modify_font","slide":1,"target":{"type":"title"},"params":{"bold":true}}]'
  python pptx_editor_llm.py <pptx文件> --exec-actions actions.json --dry-run

  # 查看结构
  python pptx_editor_llm.py <pptx文件> --inspect

环境变量 (仅 LLM 模式需要):
  OPENAI_API_KEY       API密钥 (必需)
  OPENAI_API_BASE      API端点 (可选, 兼容OpenAI的API)
  OPENAI_BASE_URL      同上 (备选)
  OPENAI_MODEL         模型名 (默认 gpt-4o)
"""

import sys, os, json, argparse, requests

# ---------------------------------------------------------------------------
# LLM 系统提示词
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = r"""你是一个 PowerPoint 编辑助手。用户会给你一个 PPTX 文件的结构信息和一条自然语言指令。
你的任务是将指令转换为一个 JSON 数组，每个元素是一个操作。

## 可用操作 (action) 及参数

### 文本操作
- modify_text: 修改文本
  {action:"modify_text", slide:1, target:{type:"title"}, params:{new_text:"新文本"}}
  target 可含: type(title/subtitle/body/table/picture), position(左上/中中/右下等), text_match(匹配文本片段)

- modify_font: 修改字体样式
  {action:"modify_font", slide:1, target:{...}, params:{font_size:24, bold:true, italic:false, underline:false, strikethrough:false, color:255, font_name:"微软雅黑", font_size_factor:1.5}}
  color 是 BGR 整数! 参数均可选，只传需要修改的。

- set_alignment: 设置对齐
  {action:"set_alignment", slide:1, target:{...}, params:{align:"center"}}
  align: left/center/right/justify 或 左/居中/右/两端

### 形状外观
- set_fill: 填充颜色
  {action:"set_fill", slide:1, target:{...}, params:{color_bgr:16711680}}

- set_border: 边框
  {action:"set_border", slide:1, target:{...}, params:{color_bgr:255, weight:2}}

### 位置/大小
- move_shape: 移动 (单位: points, 72pt = 1 inch)
  {action:"move_shape", slide:1, target:{...}, params:{left:100, top:200}}

- resize_shape: 缩放
  {action:"resize_shape", slide:1, target:{...}, params:{width:400, height:300}}
  或用 scale_factor: {action:"resize_shape", slide:1, target:{...}, params:{scale_factor:1.5}}

- delete: 删除形状
  {action:"delete", slide:1, target:{...}, params:{}}

### 添加元素
- add_textbox: 添加文本框
  {action:"add_textbox", slide:1, params:{text:"内容", left:100, top:100, width:300, height:50}}

- add_picture: 插入图片
  {action:"add_picture", slide:1, params:{pic_path:"image.png", left:100, top:100, width:200, height:150}}

### 幻灯片管理
- add_slide: 添加幻灯片
  {action:"add_slide", params:{index:3, layout:12}}
  layout: 1=标题, 2=标题+正文, 12=空白

- delete_slide: 删除幻灯片
  {action:"delete_slide", slide:2}

- move_slide: 移动幻灯片
  {action:"move_slide", slide:2, params:{new_pos:1}}

### 表格操作
- modify_cell: 修改单元格 (1-based行列)
  {action:"modify_cell", slide:1, target:{type:"table"}, params:{row:1, col:2, text:"新内容"}}

- table_row_add / table_row_delete / table_col_add / table_col_delete
  {action:"table_row_add", slide:1, target:{type:"table"}}
  {action:"table_row_delete", slide:1, target:{type:"table"}, params:{row:3}}
  {action:"table_col_add", slide:1, target:{type:"table"}}
  {action:"table_col_delete", slide:1, target:{type:"table"}, params:{col:2}}

### 动画 (COM独有)
- animation: 添加动画
  {action:"animation", slide:1, target:{...}, params:{effect:"fade"}}
  effect: appear/fly/fade/zoom/bounce

- remove_animation: 删除动画
  {action:"remove_animation", slide:1, params:{anim_index:1}}

### 切换效果
- transition: 幻灯片切换效果
  {action:"transition", slide:1, params:{transition:"fade", duration:1.5}}

### 导出
- export_pdf: 导出PDF
  {action:"export_pdf", params:{}}

- export_image: 导出幻灯片为图片
  {action:"export_image", slide:1, params:{output_path:"slide1.png", width:1920, height:1080}}

## BGR 颜色参考 (COM用BGR，不是RGB!)
红色=255(0x0000FF), 蓝色=16711680(0xFF0000), 绿色=43520(0x00AA00),
黄色=55039(0x00D7FF), 黑色=0, 白色=16777215(0xFFFFFF),
橙色=36095(0x008CFF), 紫色=8388736(0x800080), 粉色=11823615(0xB469FF),
灰色=8947848(0x888888)
RGB转BGR公式: BGR = R + G*256 + B*65536

## 输出规则
1. 只输出 JSON 数组，不要其他文字
2. slide 是 1-based 页码
3. target 用于定位已有形状; 不需要 target 时可省略
4. target.position 是形状的【当前位置】，用于定位形状，不是移动/缩放的目标位置。移动目标放在 params.left / params.top。例如"把标题移到左上"→ target:{type:"title"}, params:{left:0, top:0}
5. 若指令模糊且你无法确定，输出 {"clarify": "你的问题"} (单个对象，非数组)
6. 一条指令可能需要多个操作，全部放入数组
"""


# ---------------------------------------------------------------------------
# LLM API 调用
# ---------------------------------------------------------------------------
def get_api_config(args=None):
    """从环境变量和参数获取 API 配置"""
    base = (getattr(args, 'api_base', None) if args else None) \
        or os.environ.get('OPENAI_API_BASE') \
        or os.environ.get('OPENAI_BASE_URL') \
        or 'https://api.openai.com/v1'
    model = (getattr(args, 'model', None) if args else None) \
        or os.environ.get('OPENAI_MODEL') \
        or 'gpt-4o'
    key = (getattr(args, 'api_key', None) if args else None) \
        or os.environ.get('OPENAI_API_KEY') \
        or ''
    return base.rstrip('/'), model, key


def call_llm(messages, api_base, model, api_key, temperature=0.0):
    """调用 OpenAI-compatible chat completion API"""
    url = f"{api_base}/chat/completions"
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": 4096,
    }

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"].strip()
    except requests.exceptions.HTTPError as e:
        print(f"❌ API 错误 ({resp.status_code}): {resp.text[:500]}")
        raise
    except requests.exceptions.ConnectionError:
        print(f"❌ 无法连接 API: {url}")
        raise
    except Exception as e:
        print(f"❌ LLM 调用失败: {e}")
        raise


# ---------------------------------------------------------------------------
# 意图解析
# ---------------------------------------------------------------------------
def parse_intent_llm(instruction, pptx_structure, api_base, model, api_key,
                     conversation_history=None):
    """用 LLM 解析自然语言指令为结构化操作列表"""
    structure_text = json.dumps(pptx_structure, ensure_ascii=False, indent=2)

    messages = [{"role": "system", "content": SYSTEM_PROMPT}]

    if conversation_history:
        messages.extend(conversation_history)

    user_msg = f"## 当前 PPTX 结构\n```json\n{structure_text}\n```\n\n## 用户指令\n{instruction}"
    messages.append({"role": "user", "content": user_msg})

    raw = call_llm(messages, api_base, model, api_key)

    # 提取 JSON (LLM 可能包裹在 ```json ... ``` 中)
    text = raw.strip()
    if text.startswith("```"):
        lines = text.split("\n")
        # Remove first and last ``` lines
        start = 1
        end = len(lines)
        for i in range(len(lines) - 1, 0, -1):
            if lines[i].strip().startswith("```"):
                end = i
                break
        text = "\n".join(lines[start:end])

    try:
        result = json.loads(text)
    except json.JSONDecodeError:
        print(f"⚠️  LLM 返回非法 JSON:\n{raw[:500]}")
        return None

    return result


# ---------------------------------------------------------------------------
# 执行操作
# ---------------------------------------------------------------------------
def execute_actions(ppt, actions, dry_run=False):
    """执行 LLM 解析出的操作列表"""
    if isinstance(actions, dict):
        # 可能是 clarify 请求
        if "clarify" in actions:
            print(f"❓ LLM 需要澄清: {actions['clarify']}")
            return False
        actions = [actions]

    if not isinstance(actions, list):
        print(f"⚠️  无法识别的操作格式: {type(actions)}")
        return False

    print(f"\n📋 解析出 {len(actions)} 个操作:")
    for i, act in enumerate(actions):
        print(f"  {i+1}. {act.get('action', '?')} "
              f"(slide={act.get('slide','*')}) "
              f"params={json.dumps(act.get('params',{}), ensure_ascii=False)}")

    if dry_run:
        print("\n🔍 Dry-run 模式，不执行操作")
        return True

    print()
    for i, act in enumerate(actions):
        action = act.get("action", "")
        slide = act.get("slide")
        target = act.get("target", {})
        params = act.get("params", {})

        try:
            result = _dispatch(ppt, action, slide, target, params)
            print(f"  ✅ [{i+1}] {result}")
        except Exception as e:
            print(f"  ❌ [{i+1}] {action} 失败: {e}")

    return True


def _dispatch(ppt, action, slide, target, params):
    """分派单个操作到 PowerPointCOM 方法"""

    # --- 通用工具操作 ---
    if action == "sleep":
        import time
        seconds = params.get("seconds", 1)
        time.sleep(seconds)
        return f"等待 {seconds}s"

    # --- 幻灯片级操作 (无需 find_shape) ---
    if action == "add_slide":
        return ppt.add_slide(index=params.get("index"), layout=params.get("layout", 1))
    if action == "delete_slide":
        return ppt.delete_slide(slide)
    if action == "move_slide":
        return ppt.move_slide(slide, params["new_pos"])
    if action == "add_textbox":
        return ppt.add_textbox(slide, params["text"],
                               left=params.get("left", 100), top=params.get("top", 100),
                               width=params.get("width", 300), height=params.get("height", 50))
    if action == "add_picture":
        return ppt.add_picture(slide, params["pic_path"],
                               left=params.get("left", 100), top=params.get("top", 100),
                               width=params.get("width", 200), height=params.get("height", 150))
    if action == "export_pdf":
        pdf_path = ppt.filepath.rsplit(".", 1)[0] + ".pdf"
        ppt.prs.SaveAs(os.path.abspath(pdf_path), 32)  # ppSaveAsPDF=32
        return f"导出PDF: {pdf_path}"
    if action == "export_image":
        out = params.get("output_path", f"slide{slide}.png")
        w = params.get("width", 1920)
        h = params.get("height", 1080)
        ppt.prs.Slides(slide).Export(os.path.abspath(out), "PNG", w, h)
        return f"导出图片: {out}"
    if action == "transition":
        emap = {"fade": 1, "push": 2, "wipe": 3, "split": 4, "reveal": 5,
                "random": 6, "dissolve": 7, "checkerboard": 8, "blinds": 9, "none": 0}
        s = ppt.prs.Slides(slide)
        t = params.get("transition", "fade")
        s.SlideShowTransition.EntryEffect = emap.get(t, 1)
        if "duration" in params:
            s.SlideShowTransition.Duration = params["duration"]
        return f"第{slide}页切换效果: {t}"
    if action == "remove_animation":
        idx = params.get("anim_index", 1)
        seq = ppt.prs.Slides(slide).TimeLine.MainSequence
        if idx <= seq.Count:
            seq.Item(idx).Delete()
            return f"第{slide}页删除第{idx}个动画"
        return f"第{slide}页无第{idx}个动画"

    # --- 表格特殊操作 ---
    if action == "modify_cell":
        return ppt.modify_cell(slide, target, params["row"], params["col"], params["text"])

    if action in ("table_row_add", "table_row_delete", "table_col_add", "table_col_delete"):
        shapes = ppt.find_shape(slide, target if target else {"type": "table"})
        if not shapes:
            return f"第{slide}页未找到表格"
        shape = shapes[0]
        if action == "table_row_add":
            return ppt.add_table_row(shape)
        if action == "table_row_delete":
            return ppt.delete_table_row(shape, params["row"])
        if action == "table_col_add":
            return ppt.add_table_column(shape)
        if action == "table_col_delete":
            return ppt.delete_table_column(shape, params["col"])

    # --- 需要 find_shape 的操作 ---
    if not target:
        raise ValueError(f"操作 {action} 需要 target 来定位形状")
    shapes = ppt.find_shape(slide, target)
    if not shapes:
        raise ValueError(f"第{slide}页未找到匹配 {json.dumps(target, ensure_ascii=False)} 的形状")

    results = []
    for shape in shapes:
        if action == "modify_text":
            results.append(ppt.modify_text(shape, params["new_text"]))
        elif action == "modify_font":
            results.append(ppt.modify_font(shape, **params))
        elif action == "set_alignment":
            results.append(ppt.set_alignment(shape, params["align"]))
        elif action == "set_fill":
            results.append(ppt.set_fill(shape, params["color_bgr"]))
        elif action == "set_border":
            results.append(ppt.set_border(shape, **params))
        elif action == "move_shape":
            results.append(ppt.move_shape(shape, left=params.get("left"), top=params.get("top")))
        elif action == "resize_shape":
            if "scale_factor" in params:
                f = params["scale_factor"]
                results.append(ppt.resize_shape(shape,
                               width=shape.Width * f, height=shape.Height * f))
            else:
                results.append(ppt.resize_shape(shape,
                               width=params.get("width"), height=params.get("height")))
        elif action == "delete":
            results.append(ppt.delete_shape(shape))
            break  # 删除后不继续遍历
        elif action == "animation":
            results.append(ppt.add_animation(slide, shape, params.get("effect", "appear")))
        else:
            raise ValueError(f"未知操作: {action}")

    return "; ".join(results)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def run_single(pptx_path, instruction, output=None, dry_run=False,
               api_base=None, model=None, api_key=None, headed=False):
    """单次指令模式 (可作为模块调用)"""
    from pptx_editor_com import PowerPointCOM

    _api_base = api_base or get_api_config()[0]
    _model = model or get_api_config()[1]
    _api_key = api_key or get_api_config()[2]

    ppt = PowerPointCOM(visible=headed)
    try:
        ppt.open(pptx_path)
        structure = ppt.inspect()

        actions = parse_intent_llm(instruction, structure, _api_base, _model, _api_key)
        if actions is None:
            return False

        ok = execute_actions(ppt, actions, dry_run=dry_run)
        if ok and not dry_run:
            ppt.save(output)
        return ok
    finally:
        ppt.close()


def run_interactive(pptx_path, output=None, api_base=None, model=None, api_key=None, headed=False):
    """交互式多轮对话模式"""
    from pptx_editor_com import PowerPointCOM

    _api_base = api_base or get_api_config()[0]
    _model = model or get_api_config()[1]
    _api_key = api_key or get_api_config()[2]

    ppt = PowerPointCOM(visible=headed)
    try:
        ppt.open(pptx_path)
        structure = ppt.inspect()
        ppt.print_structure(structure)

        conversation_history = []
        print("\n💬 交互模式 (输入 quit/exit 退出, inspect 查看结构)")

        while True:
            try:
                instruction = input("\n🎯 指令> ").strip()
            except (EOFError, KeyboardInterrupt):
                break

            if not instruction:
                continue
            if instruction.lower() in ("quit", "exit", "q"):
                break
            if instruction.lower() == "inspect":
                structure = ppt.inspect()
                ppt.print_structure(structure)
                continue

            # 每次刷新结构以反映修改
            structure = ppt.inspect()

            actions = parse_intent_llm(instruction, structure, _api_base, _model, _api_key,
                                        conversation_history)
            if actions is None:
                continue

            if isinstance(actions, dict) and "clarify" in actions:
                print(f"❓ {actions['clarify']}")
                conversation_history.append({"role": "assistant", "content": json.dumps(actions, ensure_ascii=False)})
                continue

            execute_actions(ppt, actions)

            # 记录对话历史
            conversation_history.append({"role": "user", "content": instruction})
            conversation_history.append({"role": "assistant",
                                          "content": json.dumps(actions, ensure_ascii=False)})

            # 限制历史长度
            if len(conversation_history) > 20:
                conversation_history = conversation_history[-14:]

        ppt.save(output)
    finally:
        ppt.close()


def main():
    parser = argparse.ArgumentParser(description="PPTX 自然语言编辑器 (LLM + COM)")
    parser.add_argument("pptx_file", help="PPTX 文件路径")
    parser.add_argument("instruction", nargs="?", default=None, help="编辑指令 (支持中英文)")
    parser.add_argument("--interactive", "-i", action="store_true", help="交互式多轮对话模式")
    parser.add_argument("--inspect", action="store_true", help="查看 PPTX 结构")
    parser.add_argument("--output", "-o", help="输出文件路径")
    parser.add_argument("--dry-run", action="store_true", help="仅解析意图，不执行")
    parser.add_argument("--exec-script", metavar="SCRIPT", help="执行Python脚本文件 (脚本内可用 ppt, filepath 变量, 无需API)")
    parser.add_argument("--exec-actions", metavar="JSON", help="执行JSON动作数组 (字符串或.json文件路径, 无需API)")
    parser.add_argument("--api-base", help="API 端点")
    parser.add_argument("--model", help="模型名称")
    parser.add_argument("--api-key", help="API 密钥")
    parser.add_argument("--headed", action="store_true", help="以可见窗口模式打开 PowerPoint")
    args = parser.parse_args()

    if not os.path.exists(args.pptx_file):
        print(f"❌ 文件不存在: {args.pptx_file}")
        sys.exit(1)

    api_base, model, api_key = get_api_config(args)

    if args.inspect:
        from pptx_editor_com import PowerPointCOM
        ppt = PowerPointCOM(visible=args.headed)
        try:
            ppt.open(args.pptx_file)
            structure = ppt.inspect()
            ppt.print_structure(structure)
            print("\n" + json.dumps(structure, ensure_ascii=False, indent=2))
        finally:
            ppt.close()
        if not args.exec_script and not args.exec_actions:
            return

    if args.exec_script:
        from pptx_editor_com import PowerPointCOM
        ppt = PowerPointCOM(visible=args.headed)
        try:
            ppt.open(args.pptx_file)
            if args.inspect and not args.exec_actions:
                pass  # already printed above
            script_path = args.exec_script
            if not os.path.exists(script_path):
                print(f"❌ 脚本文件不存在: {script_path}")
                sys.exit(1)
            print(f"🔧 执行脚本: {script_path}")
            with open(script_path, "r", encoding="utf-8") as f:
                script_code = f.read()
            exec(script_code, {"ppt": ppt, "filepath": os.path.abspath(args.pptx_file),
                               "__builtins__": __builtins__})
            ppt.save(args.output)
            print("✅ 脚本执行完成")
        except Exception as e:
            print(f"❌ 脚本执行失败: {e}")
            import traceback; traceback.print_exc()
            sys.exit(1)
        finally:
            ppt.close()
        return

    if args.exec_actions:
        from pptx_editor_com import PowerPointCOM
        raw = args.exec_actions
        # Try as file path first
        if os.path.exists(raw):
            with open(raw, "r", encoding="utf-8") as f:
                raw = f.read()
        try:
            actions = json.loads(raw)
        except json.JSONDecodeError as e:
            print(f"❌ JSON 解析失败: {e}")
            sys.exit(1)
        ppt = PowerPointCOM(visible=args.headed)
        try:
            ppt.open(args.pptx_file)
            if args.inspect:
                pass  # already printed above
            execute_actions(ppt, actions, dry_run=args.dry_run)
            if not args.dry_run:
                ppt.save(args.output)
        except Exception as e:
            print(f"❌ 动作执行失败: {e}")
            import traceback; traceback.print_exc()
            sys.exit(1)
        finally:
            ppt.close()
        return

    if args.interactive:
        run_interactive(args.pptx_file, args.output, api_base, model, api_key, headed=args.headed)
        return

    if not args.instruction:
        # 尝试从 stdin 读取批量指令
        if not sys.stdin.isatty():
            instructions = sys.stdin.read().strip()
        else:
            print("❌ 请提供指令或使用 --interactive 模式")
            parser.print_help()
            sys.exit(1)
    else:
        instructions = args.instruction

    # 批量模式: 换行分隔多条指令
    lines = [l.strip() for l in instructions.split("\n") if l.strip()]

    from pptx_editor_com import PowerPointCOM
    ppt = PowerPointCOM(visible=args.headed)
    try:
        ppt.open(args.pptx_file)

        for idx, instruction in enumerate(lines):
            print(f"\n{'='*50}")
            print(f"📝 指令 [{idx+1}/{len(lines)}]: {instruction}")
            print(f"{'='*50}")

            structure = ppt.inspect()
            actions = parse_intent_llm(instruction, structure, api_base, model, api_key)
            if actions is None:
                continue
            execute_actions(ppt, actions, dry_run=args.dry_run)

        if not args.dry_run:
            ppt.save(args.output)
    finally:
        ppt.close()


if __name__ == "__main__":
    main()
