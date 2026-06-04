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


CLOSE_DELAY_SECONDS = 30


def _resolve_note_slide(ppt, slide=None, note_slide=None):
    """Choose which slide should receive progress notes."""
    if note_slide is not None:
        return max(1, min(note_slide, ppt.prs.Slides.Count))
    if isinstance(slide, int) and 1 <= slide <= ppt.prs.Slides.Count:
        return slide
    return 1


def _write_progress_note(ppt, message, slide=None, note_slide=None, append=False):
    """Write or append progress text to speaker notes."""
    target_slide = _resolve_note_slide(ppt, slide=slide, note_slide=note_slide)
    if append:
        return ppt.append_notes(target_slide, message)
    return ppt.set_notes(target_slide, message)


def _format_progress_entry(detail, slide=None, index=None, success=True):
    """Format one incremental notes entry for an executed action."""
    prefix = "✅" if success else "❌"
    text = detail
    if isinstance(slide, int) and slide > 0 and f"第{slide}页" not in detail:
        text = f"第{slide}页: {detail}"
    if index is not None:
        return f"{prefix} [{index}] {text}"
    return f"{prefix} {text}"


def _build_script_helpers(ppt, note_slide=None):
    """Expose note logging and sleep helpers inside --exec-script mode."""
    def log_note(message, slide=None, append=True):
        return _write_progress_note(
            ppt,
            message,
            slide=slide,
            note_slide=note_slide,
            append=append,
        )

    def sleep(seconds, slide=None, note=None, append=True):
        import time

        if note:
            log_note(note, slide=slide, append=append)
        time.sleep(seconds)
        return f"等待 {seconds}s"

    return log_note, sleep


def _close_after_grace(ppt, modified=False):
    """Keep PowerPoint open briefly after real modifications."""
    if modified:
        import time

        print(f"⏳ 已完成修改，保留 PowerPoint {CLOSE_DELAY_SECONDS}s 后关闭...")
        time.sleep(CLOSE_DELAY_SECONDS)
    ppt.close()


def _load_actions_input(raw):
    """Load JSON actions from a raw string or file path."""
    text = raw
    if os.path.exists(raw):
        with open(raw, "r", encoding="utf-8") as f:
            text = f.read()
    return json.loads(text)


def _run_script_code(ppt, script_code, script_label, note_slide=None, notes_progress=False):
    """Execute script code with the injected helper globals."""
    log_note, sleep = _build_script_helpers(ppt, note_slide=note_slide)
    if notes_progress:
        _write_progress_note(
            ppt,
            f"开始执行脚本: {script_label}",
            note_slide=note_slide,
            append=True,
        )
    exec(script_code, {
        "ppt": ppt,
        "filepath": ppt.filepath,
        "log_note": log_note,
        "sleep": sleep,
        "__builtins__": __builtins__,
    })
    if notes_progress:
        _write_progress_note(
            ppt,
            f"脚本执行完成: {script_label}",
            note_slide=note_slide,
            append=True,
        )


def _run_script_path(ppt, script_path, note_slide=None, notes_progress=False):
    """Execute a script file by path."""
    if not os.path.exists(script_path):
        raise FileNotFoundError(f"脚本文件不存在: {script_path}")
    with open(script_path, "r", encoding="utf-8") as f:
        script_code = f.read()
    _run_script_code(
        ppt,
        script_code,
        os.path.basename(script_path),
        note_slide=note_slide,
        notes_progress=notes_progress,
    )


def _print_session_help(mode):
    """Print available commands for interactive action/script sessions."""
    print("\n💡 会话命令:")
    print("  inspect [slide]         查看当前结构或指定页结构")
    print("  status                  查看当前会话状态")
    print("  save                    保存当前文件")
    print("  saveas <path>           另存为指定文件")
    print("  quit                    保存并退出")
    print("  close                   直接关闭不再继续读取命令")
    print("  help                    显示本帮助")
    if mode == "actions":
        print("  {\"command\":\"actions\",\"payload\":[...]}  执行 JSON actions")
        print("  [...]/{...}             直接执行 JSON actions")
    else:
        print("  {\"command\":\"script\",\"path\":\"edit.py\"}        执行脚本文件")
        print("  {\"command\":\"script_inline\",\"code\":\"...\"}  执行内联脚本")
        print("  <script.py>             直接执行脚本文件")


def _print_session_status(ppt, mode, output_path, modified):
    """Print interactive session status."""
    print(json.dumps({
        "mode": mode,
        "file": ppt.filepath,
        "slides": ppt.prs.Slides.Count,
        "output": output_path or ppt.filepath,
        "modified": modified,
    }, ensure_ascii=False, indent=2))


def _parse_session_command(raw, mode):
    """Parse a line of interactive session input."""
    text = raw.strip()
    if not text:
        return None

    lowered = text.lower()
    if lowered == "inspect" or lowered.startswith("inspect "):
        parts = text.split(maxsplit=1)
        slide = None
        if len(parts) > 1:
            try:
                slide = int(parts[1].strip())
            except ValueError as exc:
                raise ValueError("inspect 后只支持页码，例如 inspect 3") from exc
            if slide < 1:
                raise ValueError("inspect 页码必须从 1 开始")
        return {"command": "inspect", "slide": slide}
    if lowered in ("help", "status", "save", "quit", "exit", "close"):
        return {"command": lowered}
    if lowered.startswith("saveas "):
        return {"command": "saveas", "path": text[7:].strip()}

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        if mode == "script":
            return {"command": "script", "path": text}
        raise ValueError("无法解析输入。请提供 JSON 命令、脚本路径或 help。")

    if isinstance(parsed, dict) and "command" in parsed:
        return parsed

    if mode == "actions":
        if isinstance(parsed, list):
            return {"command": "actions", "payload": parsed}
        if isinstance(parsed, dict) and ("action" in parsed or "clarify" in parsed):
            return {"command": "actions", "payload": parsed}

    raise ValueError("输入格式不受支持。请使用 help 查看协议。")


def _run_interactive_session(pptx_path, mode, initial_payload=None, output=None,
                             dry_run=False, headed=False, notes_progress=False,
                             note_slide=None):
    """Keep one COM session open and execute commands from stdin."""
    from pptx_editor_com import PowerPointCOM

    ppt = PowerPointCOM(visible=headed)
    modified = False
    current_output = output

    try:
        ppt.open(pptx_path)
        print(f"\n💬 {mode} 会话模式 (输入 help 查看命令, quit 保存退出)")

        if initial_payload and initial_payload != "-":
            if mode == "actions":
                actions = _load_actions_input(initial_payload)
                if notes_progress and not dry_run:
                    _write_progress_note(
                        ppt,
                        f"执行动作集: {len(actions) if isinstance(actions, list) else 1} 个动作",
                        note_slide=note_slide,
                        append=True,
                    )
                execute_actions(
                    ppt,
                    actions,
                    dry_run=dry_run,
                    progress_callback=(
                        lambda message, slide=None: _write_progress_note(
                            ppt,
                            message,
                            slide=slide,
                            note_slide=note_slide,
                            append=True,
                        )
                    ) if notes_progress and not dry_run else None,
                )
                if not dry_run:
                    modified = True
            else:
                _run_script_path(ppt, initial_payload, note_slide=note_slide, notes_progress=notes_progress)
                modified = True

        while True:
            try:
                raw = input(f"\n{mode}> ")
            except EOFError:
                raw = "quit"
            except KeyboardInterrupt:
                raw = "quit"
                print()

            try:
                command = _parse_session_command(raw, mode)
            except Exception as e:
                print(f"❌ {e}")
                continue

            if command is None:
                continue

            cmd = command.get("command", "").lower()

            if cmd == "help":
                _print_session_help(mode)
                continue
            if cmd == "inspect":
                structure = ppt.inspect()
                slide = command.get("slide")
                if slide is None:
                    ppt.print_structure(structure)
                    continue
                slide_data = [item for item in structure["slides"] if item.get("index") == slide]
                if not slide_data:
                    print(f"❌ 第 {slide} 页不存在")
                    continue
                ppt.print_structure({"slides": slide_data})
                continue
            if cmd == "status":
                _print_session_status(ppt, mode, current_output, modified)
                continue
            if cmd == "save":
                if dry_run:
                    print("🔍 Dry-run 模式，未保存")
                else:
                    ppt.save(current_output)
                    print("✅ 已保存")
                continue
            if cmd == "saveas":
                saveas_path = command.get("path")
                if not saveas_path:
                    print("❌ saveas 需要 path")
                    continue
                if dry_run:
                    print(f"🔍 Dry-run 模式，未另存为: {saveas_path}")
                else:
                    current_output = saveas_path
                    ppt.save(current_output)
                    print(f"✅ 已另存为: {current_output}")
                continue
            if cmd in ("quit", "exit"):
                if not dry_run:
                    ppt.save(current_output)
                break
            if cmd == "close":
                modified = False
                break

            try:
                if mode == "actions" and cmd == "actions":
                    payload = command.get("payload")
                    if notes_progress and not dry_run:
                        count = len(payload) if isinstance(payload, list) else 1
                        _write_progress_note(
                            ppt,
                            f"执行动作集: {count} 个动作",
                            note_slide=note_slide,
                            append=True,
                        )
                    execute_actions(
                        ppt,
                        payload,
                        dry_run=dry_run,
                        progress_callback=(
                            lambda message, slide=None: _write_progress_note(
                                ppt,
                                message,
                                slide=slide,
                                note_slide=note_slide,
                                append=True,
                            )
                        ) if notes_progress and not dry_run else None,
                    )
                    if notes_progress and not dry_run:
                        count = len(payload) if isinstance(payload, list) else 1
                        _write_progress_note(
                            ppt,
                            f"完成动作集: {count} 个动作",
                            note_slide=note_slide,
                            append=True,
                        )
                    if not dry_run:
                        modified = True
                    continue

                if mode == "script" and cmd == "script":
                    _run_script_path(
                        ppt,
                        command.get("path", ""),
                        note_slide=note_slide,
                        notes_progress=notes_progress,
                    )
                    modified = True
                    continue

                if mode == "script" and cmd == "script_inline":
                    _run_script_code(
                        ppt,
                        command.get("code", ""),
                        "<inline>",
                        note_slide=note_slide,
                        notes_progress=notes_progress,
                    )
                    modified = True
                    continue

                print("❌ 不支持的命令，输入 help 查看用法")
            except Exception as e:
                print(f"❌ 执行失败: {e}")

    finally:
        _close_after_grace(ppt, modified=modified and not dry_run)

# ---------------------------------------------------------------------------
# LLM 系统提示词
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = r"""你是一个 PowerPoint 编辑助手。用户会给你一个 PPTX 文件的结构信息和一条自然语言指令。
你的任务是将指令转换为一个 JSON 数组，每个元素是一个操作。

## 可用操作 (action) 及参数

### 文本操作
- modify_text: 修改文本
  {action:"modify_text", slide:1, target:{type:"title"}, params:{new_text:"新文本"}}
    target 可含: type(title/subtitle/body/table/picture/chart/textbox), position(左上/中中/右下等), text_match(匹配文本片段)

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

- add_table: 添加表格
    {action:"add_table", slide:1, params:{rows:3, cols:4, left:100, top:100, width:400, height:200}}

- add_chart: 添加图表
    {action:"add_chart", slide:1, params:{chart_type:4, data:[[1,2,3],[4,5,6]], left:100, top:100, width:400, height:300}}

- add_smartart: 添加 SmartArt
    {action:"add_smartart", slide:1, params:{layout_id:1, left:100, top:100, width:400, height:300}}

- add_audio / add_video: 插入媒体
    {action:"add_audio", slide:1, params:{audio_path:"sound.mp3", left:100, top:100, width:50, height:50}}
    {action:"add_video", slide:1, params:{video_path:"demo.mp4", left:100, top:100, width:400, height:300}}

- add_freeform: 添加自由形状
    {action:"add_freeform", slide:1, params:{points:[[100,100],[200,100],[180,180],[100,200]]}}

### 幻灯片管理
- add_slide: 添加幻灯片
  {action:"add_slide", params:{index:3, layout:12}}
  layout: 1=标题, 2=标题+正文, 12=空白

- delete_slide: 删除幻灯片
  {action:"delete_slide", slide:2}

- move_slide: 移动幻灯片
  {action:"move_slide", slide:2, params:{new_pos:1}}

- duplicate_slide: 复制幻灯片
    {action:"duplicate_slide", slide:2}

- set_slide_size / set_slide_size_preset: 设定页面尺寸
    {action:"set_slide_size", params:{width:960, height:540}}
    {action:"set_slide_size_preset", params:{preset:"widescreen"}}

- set_slide_background / set_slide_background_image: 设置背景
    {action:"set_slide_background", slide:1, params:{color_bgr:16777215}}
    {action:"set_slide_background_image", slide:1, params:{image_path:"bg.png"}}

- set_notes / append_notes: 设置备注
    {action:"set_notes", slide:1, params:{text:"演讲者备注"}}
    {action:"append_notes", slide:1, params:{text:"补充备注", separator:"\n"}}

- add_comment / delete_comment: 幻灯片评论
    {action:"add_comment", slide:1, params:{text:"需要复核", author:"Reviewer", x:10, y:10}}
    {action:"delete_comment", slide:1, params:{comment_idx:1}}

- add_section / delete_section / rename_section: 分节管理
    {action:"add_section", params:{name:"Overview", slide_idx:1}}
    {action:"delete_section", params:{section_idx:1}}
    {action:"rename_section", params:{section_idx:1, new_name:"Intro"}}

- set_slideshow_settings / start_slideshow: 放映设置
    {action:"set_slideshow_settings", params:{loop:false, show_narration:true, show_animation:true}}
    {action:"start_slideshow", params:{from_slide:1, to_slide:3}}

- merge_presentations / print_presentation: 合并或打印
    {action:"merge_presentations", params:{file_paths:["other.pptx"], output_path:"merged.pptx"}}
    {action:"print_presentation", params:{printer_name:"Microsoft Print to PDF", copies:1}}

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

- modify_animation_effect: 修改动画效果
    {action:"modify_animation_effect", slide:1, params:{anim_index:1, effect:"zoom"}}

### 形状增强
- rotate_shape / flip_shape / set_zorder
    {action:"rotate_shape", slide:1, target:{...}, params:{angle:45}}
    {action:"flip_shape", slide:1, target:{...}, params:{direction:"horizontal"}}
    {action:"set_zorder", slide:1, target:{...}, params:{position:"front"}}

- crop_picture / set_brightness / set_contrast / replace_picture
    {action:"crop_picture", slide:1, target:{type:"picture"}, params:{left:10, top:0, right:10, bottom:0}}
    {action:"set_brightness", slide:1, target:{type:"picture"}, params:{value:0.4}}
    {action:"set_contrast", slide:1, target:{type:"picture"}, params:{value:0.3}}
    {action:"replace_picture", slide:1, target:{type:"picture"}, params:{new_path:"new.png"}}

- add_bullet / set_text_autofit / add_hyperlink / set_word_art
    {action:"add_bullet", slide:1, target:{...}, params:{level:1}}
    {action:"set_text_autofit", slide:1, target:{...}, params:{mode:"fit"}}
    {action:"add_hyperlink", slide:1, target:{...}, params:{url:"https://example.com", text:"访问链接"}}
    {action:"set_word_art", slide:1, target:{...}, params:{style:1}}

- set_line_spacing / set_paragraph_spacing
    {action:"set_line_spacing", slide:1, target:{...}, params:{spacing:1.5}}
    {action:"set_paragraph_spacing", slide:1, target:{...}, params:{before:6, after:4}}

- set_shadow / set_reflection / set_glow / set_3d_rotation
    {action:"set_shadow", slide:1, target:{...}, params:{preset:1}}
    {action:"set_reflection", slide:1, target:{...}, params:{preset:1}}
    {action:"set_glow", slide:1, target:{...}, params:{color_bgr:255, radius:10}}
    {action:"set_3d_rotation", slide:1, target:{...}, params:{x:10, y:20, z:0}}

- set_media_playback
    {action:"set_media_playback", slide:1, target:{name:"Video 1"}, params:{auto_play:true, loop:false, hide_on_stop:false}}

- set_chart_title / set_chart_style / modify_chart_data
    {action:"set_chart_title", slide:1, target:{type:"chart"}, params:{title:"季度营收"}}
    {action:"set_chart_style", slide:1, target:{type:"chart"}, params:{style_id:10}}
    {action:"modify_chart_data", slide:1, target:{type:"chart"}, params:{series_idx:1, values:[1,2,3,4]}}

### 切换效果
- transition: 幻灯片切换效果
  {action:"transition", slide:1, params:{transition:"fade", duration:1.5}}

### 导出
- export_pdf: 导出PDF
  {action:"export_pdf", params:{}}

- export_image: 导出幻灯片为图片
  {action:"export_image", slide:1, params:{output_path:"slide1.png", width:1920, height:1080}}

### 工具动作
- sleep: 等待若干秒，便于 headed 模式下观察修改过程
    {action:"sleep", params:{seconds:2}}

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
def execute_actions(ppt, actions, dry_run=False, progress_callback=None):
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
            entry = _format_progress_entry(result, slide=slide, index=i + 1, success=True)
            print(f"  {entry}")
            if progress_callback:
                progress_callback(entry, slide=slide)
        except Exception as e:
            entry = _format_progress_entry(f"{action} 失败: {e}", slide=slide, index=i + 1, success=False)
            print(f"  {entry}")
            if progress_callback:
                progress_callback(entry, slide=slide)

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
    if action == "add_table":
        return ppt.add_table(slide, params["rows"], params["cols"],
                             left=params.get("left", 100), top=params.get("top", 100),
                             width=params.get("width", 400), height=params.get("height", 200))
    if action == "add_chart":
        return ppt.add_chart(slide, params.get("chart_type", 4), params.get("data"),
                             left=params.get("left", 100), top=params.get("top", 100),
                             width=params.get("width", 400), height=params.get("height", 300))
    if action == "add_smartart":
        return ppt.add_smartart(slide, params.get("layout_id"),
                                left=params.get("left", 100), top=params.get("top", 100),
                                width=params.get("width", 400), height=params.get("height", 300))
    if action == "add_audio":
        return ppt.add_audio(slide, params["audio_path"],
                             left=params.get("left", 100), top=params.get("top", 100),
                             width=params.get("width", 50), height=params.get("height", 50))
    if action == "add_video":
        return ppt.add_video(slide, params["video_path"],
                             left=params.get("left", 100), top=params.get("top", 100),
                             width=params.get("width", 400), height=params.get("height", 300))
    if action == "add_freeform":
        return ppt.add_freeform(slide, params["points"])
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
    if action == "modify_animation_effect":
        return ppt.modify_animation_effect(slide, params["anim_index"], params["effect"])
    if action == "duplicate_slide":
        return ppt.duplicate_slide(slide)
    if action == "set_slide_size":
        return ppt.set_slide_size(params["width"], params["height"])
    if action == "set_slide_size_preset":
        return ppt.set_slide_size_preset(params["preset"])
    if action == "set_slide_background":
        return ppt.set_slide_background(slide, params["color_bgr"])
    if action == "set_slide_background_image":
        return ppt.set_slide_background_image(slide, params["image_path"])
    if action == "set_notes":
        return ppt.set_notes(slide, params["text"])
    if action == "append_notes":
        return ppt.append_notes(slide, params["text"], params.get("separator", "\n"))
    if action == "add_comment":
        return ppt.add_comment(slide, params["text"], params.get("author", "Author"),
                               params.get("x", 10), params.get("y", 10))
    if action == "delete_comment":
        return ppt.delete_comment(slide, params["comment_idx"])
    if action == "add_section":
        return ppt.add_section(params["name"], params["slide_idx"])
    if action == "delete_section":
        return ppt.delete_section(params["section_idx"])
    if action == "rename_section":
        return ppt.rename_section(params["section_idx"], params["new_name"])
    if action == "set_slideshow_settings":
        return ppt.set_slideshow_settings(loop=params.get("loop", False),
                                          show_narration=params.get("show_narration", True),
                                          show_animation=params.get("show_animation", True))
    if action == "start_slideshow":
        return ppt.start_slideshow(from_slide=params.get("from_slide", 1),
                                   to_slide=params.get("to_slide"))
    if action == "merge_presentations":
        return ppt.merge_presentations(params["file_paths"], params.get("output_path"))
    if action == "print_presentation":
        return ppt.print_presentation(printer_name=params.get("printer_name"),
                                      copies=params.get("copies", 1),
                                      print_range=params.get("print_range"))

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
        elif action == "rotate_shape":
            results.append(ppt.rotate_shape(shape, params["angle"]))
        elif action == "flip_shape":
            results.append(ppt.flip_shape(shape, params.get("direction", "horizontal")))
        elif action == "set_zorder":
            results.append(ppt.set_zorder(shape, params["position"]))
        elif action == "crop_picture":
            results.append(ppt.crop_picture(shape,
                           left=params.get("left", 0), top=params.get("top", 0),
                           right=params.get("right", 0), bottom=params.get("bottom", 0)))
        elif action == "set_brightness":
            results.append(ppt.set_brightness(shape, params["value"]))
        elif action == "set_contrast":
            results.append(ppt.set_contrast(shape, params["value"]))
        elif action == "replace_picture":
            results.append(ppt.replace_picture(shape, params["new_path"]))
        elif action == "add_bullet":
            results.append(ppt.add_bullet(shape, params.get("level", 1)))
        elif action == "set_text_autofit":
            results.append(ppt.set_text_autofit(shape, params["mode"]))
        elif action == "add_hyperlink":
            results.append(ppt.add_hyperlink(shape, params["url"], params.get("text")))
        elif action == "set_word_art":
            results.append(ppt.set_word_art(shape, params["style"]))
        elif action == "set_line_spacing":
            results.append(ppt.set_line_spacing(shape, params["spacing"]))
        elif action == "set_paragraph_spacing":
            results.append(ppt.set_paragraph_spacing(shape,
                           before=params.get("before", 0), after=params.get("after", 0)))
        elif action == "set_shadow":
            results.append(ppt.set_shadow(shape, params["preset"]))
        elif action == "set_reflection":
            results.append(ppt.set_reflection(shape, params["preset"]))
        elif action == "set_glow":
            results.append(ppt.set_glow(shape, params["color_bgr"], params.get("radius", 10)))
        elif action == "set_3d_rotation":
            results.append(ppt.set_3d_rotation(shape,
                           x=params.get("x", 0), y=params.get("y", 0), z=params.get("z", 0)))
        elif action == "set_media_playback":
            results.append(ppt.set_media_playback(shape,
                           auto_play=params.get("auto_play", False),
                           loop=params.get("loop", False),
                           hide_on_stop=params.get("hide_on_stop", False)))
        elif action == "modify_chart_data":
            results.append(ppt.modify_chart_data(shape, params["series_idx"], params["values"]))
        elif action == "set_chart_title":
            results.append(ppt.set_chart_title(shape, params["title"]))
        elif action == "set_chart_style":
            results.append(ppt.set_chart_style(shape, params["style_id"]))
        else:
            raise ValueError(f"未知操作: {action}")

    return "; ".join(results)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def run_single(pptx_path, instruction, output=None, dry_run=False,
               api_base=None, model=None, api_key=None,
               headed=False, notes_progress=False, note_slide=None):
    """单次指令模式 (可作为模块调用)"""
    from pptx_editor_com import PowerPointCOM

    _api_base = api_base or get_api_config()[0]
    _model = model or get_api_config()[1]
    _api_key = api_key or get_api_config()[2]

    ppt = PowerPointCOM(visible=headed)
    modified = False
    try:
        ppt.open(pptx_path)
        structure = ppt.inspect()

        actions = parse_intent_llm(instruction, structure, _api_base, _model, _api_key)
        if actions is None:
            return False

        if notes_progress and not dry_run:
            _write_progress_note(
                ppt,
                f"执行指令: {instruction}",
                note_slide=note_slide,
                append=True,
            )

        ok = execute_actions(
            ppt,
            actions,
            dry_run=dry_run,
            progress_callback=(
                lambda message, slide=None: _write_progress_note(
                    ppt,
                    message,
                    slide=slide,
                    note_slide=note_slide,
                    append=True,
                )
            ) if notes_progress and not dry_run else None,
        )
        if ok and notes_progress and not dry_run:
            _write_progress_note(
                ppt,
                f"完成指令: {instruction}",
                note_slide=note_slide,
                append=True,
            )
        if ok and not dry_run:
            ppt.save(output)
            modified = True
        return ok
    finally:
        _close_after_grace(ppt, modified=modified)


def run_interactive(pptx_path, output=None, api_base=None, model=None, api_key=None,
                    headed=False, notes_progress=False, note_slide=None):
    """交互式多轮对话模式"""
    from pptx_editor_com import PowerPointCOM

    _api_base = api_base or get_api_config()[0]
    _model = model or get_api_config()[1]
    _api_key = api_key or get_api_config()[2]

    ppt = PowerPointCOM(visible=headed)
    modified = False
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

            if notes_progress:
                _write_progress_note(
                    ppt,
                    f"执行指令: {instruction}",
                    note_slide=note_slide,
                    append=True,
                )

            execute_actions(
                ppt,
                actions,
                progress_callback=(
                    lambda message, slide=None: _write_progress_note(
                        ppt,
                        message,
                        slide=slide,
                        note_slide=note_slide,
                        append=True,
                    )
                ) if notes_progress else None,
            )
            modified = True

            if notes_progress:
                _write_progress_note(
                    ppt,
                    f"完成指令: {instruction}",
                    note_slide=note_slide,
                    append=True,
                )

            # 记录对话历史
            conversation_history.append({"role": "user", "content": instruction})
            conversation_history.append({"role": "assistant",
                                          "content": json.dumps(actions, ensure_ascii=False)})

            # 限制历史长度
            if len(conversation_history) > 20:
                conversation_history = conversation_history[-14:]

        ppt.save(output)
    finally:
        _close_after_grace(ppt, modified=modified)


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
    parser.add_argument("--interactive-actions", metavar="JSON", nargs="?", const="-", help="保持同一个 COM 会话，持续从 stdin 读取 JSON 动作命令")
    parser.add_argument("--interactive-script", metavar="SCRIPT", nargs="?", const="-", help="保持同一个 COM 会话，持续从 stdin 读取脚本命令")
    parser.add_argument("--api-base", help="API 端点")
    parser.add_argument("--model", help="模型名称")
    parser.add_argument("--api-key", help="API 密钥")
    parser.add_argument("--headed", action="store_true", help="以可见窗口模式打开 PowerPoint")
    parser.add_argument("--notes-progress", action="store_true", help="执行时将当前指令写入演讲者备注")
    parser.add_argument("--note-slide", type=int, help="将进度备注固定写到指定页")
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
            _close_after_grace(ppt, modified=False)
        if not args.exec_script and not args.exec_actions:
            return

    if args.interactive_actions is not None:
        _run_interactive_session(
            args.pptx_file,
            "actions",
            initial_payload=args.interactive_actions,
            output=args.output,
            dry_run=args.dry_run,
            headed=args.headed,
            notes_progress=args.notes_progress,
            note_slide=args.note_slide,
        )
        return

    if args.interactive_script is not None:
        _run_interactive_session(
            args.pptx_file,
            "script",
            initial_payload=args.interactive_script,
            output=args.output,
            dry_run=args.dry_run,
            headed=args.headed,
            notes_progress=args.notes_progress,
            note_slide=args.note_slide,
        )
        return

    if args.exec_script:
        from pptx_editor_com import PowerPointCOM
        ppt = PowerPointCOM(visible=args.headed)
        modified = False
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
            log_note, sleep = _build_script_helpers(ppt, note_slide=args.note_slide)
            if args.notes_progress:
                _write_progress_note(
                    ppt,
                    f"开始执行脚本: {os.path.basename(script_path)}",
                    note_slide=args.note_slide,
                    append=True,
                )
            exec(script_code, {
                "ppt": ppt,
                "filepath": os.path.abspath(args.pptx_file),
                "log_note": log_note,
                "sleep": sleep,
                "__builtins__": __builtins__,
            })
            if args.notes_progress:
                _write_progress_note(
                    ppt,
                    f"脚本执行完成: {os.path.basename(script_path)}",
                    note_slide=args.note_slide,
                    append=True,
                )
            ppt.save(args.output)
            modified = True
            print("✅ 脚本执行完成")
        except Exception as e:
            print(f"❌ 脚本执行失败: {e}")
            import traceback; traceback.print_exc()
            sys.exit(1)
        finally:
            _close_after_grace(ppt, modified=modified)
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
        modified = False
        try:
            ppt.open(args.pptx_file)
            if args.inspect:
                pass  # already printed above
            if args.notes_progress and not args.dry_run:
                _write_progress_note(
                    ppt,
                    f"执行动作集: {len(actions)} 个动作",
                    note_slide=args.note_slide,
                    append=True,
                )
            execute_actions(
                ppt,
                actions,
                dry_run=args.dry_run,
                progress_callback=(
                    lambda message, slide=None: _write_progress_note(
                        ppt,
                        message,
                        slide=slide,
                        note_slide=args.note_slide,
                        append=True,
                    )
                ) if args.notes_progress and not args.dry_run else None,
            )
            if args.notes_progress and not args.dry_run:
                _write_progress_note(
                    ppt,
                    f"完成动作集: {len(actions)} 个动作",
                    note_slide=args.note_slide,
                    append=True,
                )
            if not args.dry_run:
                ppt.save(args.output)
                modified = True
        except Exception as e:
            print(f"❌ 动作执行失败: {e}")
            import traceback; traceback.print_exc()
            sys.exit(1)
        finally:
            _close_after_grace(ppt, modified=modified)
        return

    if args.interactive:
        run_interactive(
            args.pptx_file,
            args.output,
            api_base,
            model,
            api_key,
            headed=args.headed,
            notes_progress=args.notes_progress,
            note_slide=args.note_slide,
        )
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
    modified = False
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
            if args.notes_progress and not args.dry_run:
                _write_progress_note(
                    ppt,
                    f"执行指令: {instruction}",
                    note_slide=args.note_slide,
                    append=True,
                )
            execute_actions(
                ppt,
                actions,
                dry_run=args.dry_run,
                progress_callback=(
                    lambda message, slide=None: _write_progress_note(
                        ppt,
                        message,
                        slide=slide,
                        note_slide=args.note_slide,
                        append=True,
                    )
                ) if args.notes_progress and not args.dry_run else None,
            )
            if args.notes_progress and not args.dry_run:
                _write_progress_note(
                    ppt,
                    f"完成指令: {instruction}",
                    note_slide=args.note_slide,
                    append=True,
                )
            if not args.dry_run:
                modified = True

        if not args.dry_run:
            ppt.save(args.output)
    finally:
        _close_after_grace(ppt, modified=modified)


if __name__ == "__main__":
    main()
