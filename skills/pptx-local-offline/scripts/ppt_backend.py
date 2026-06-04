import json
import os
import sys

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None


SUPPORTED_BACKENDS = ("pywin32", "vba")


def add_backend_arguments(parser):
    parser.add_argument(
        "--backend",
        choices=SUPPORTED_BACKENDS,
        default=os.environ.get("PPTX_EDITOR_BACKEND", "pywin32"),
        help="底层执行策略: pywin32(默认) 或 vba",
    )
    parser.add_argument(
        "--vba-module",
        default=os.environ.get("PPTX_EDITOR_VBA_MODULE", "PptEditorBridge"),
        help="VBA backend 使用的宏模块名，默认 PptEditorBridge",
    )


def create_backend(name="pywin32", visible=False, vba_module="PptEditorBridge"):
    backend = (name or "pywin32").lower()
    if backend == "pywin32":
        from pptx_editor_com import PowerPointCOM

        return PowerPointCOM(visible=visible)
    if backend == "vba":
        return PowerPointVBA(visible=visible, macro_module=vba_module)
    raise ValueError(f"不支持的 backend: {name}")


class PowerPointVBA:
    """PowerPoint backend that routes inspect/actions through Application.Run."""

    def __init__(self, visible=False, macro_module="PptEditorBridge"):
        if win32com is None or pythoncom is None:
            print("❌ pip install pywin32 (Windows + Office required)")
            sys.exit(1)
        pythoncom.CoInitialize()
        self.app = win32com.client.Dispatch("PowerPoint.Application")
        if visible:
            self.app.Visible = True
        self.prs = None
        self.filepath = None
        self.macro_module = macro_module

    def _run_macro(self, macro_name, *args):
        try:
            return self.app.Run(f"{self.macro_module}.{macro_name}", *args)
        except Exception as exc:
            raise RuntimeError(
                f"VBA backend 调用失败: {self.macro_module}.{macro_name}。"
                f"请先导入/实现 VBA 桥接模块。原始错误: {exc}"
            ) from exc

    def open(self, path):
        self.filepath = os.path.abspath(path)
        with_window = self.app.Visible
        self.prs = self.app.Presentations.Open(self.filepath, False, False, with_window)
        print(f"📂 已打开: {self.filepath} ({self.prs.Slides.Count}页)")
        return self

    def close(self):
        try:
            if self.prs:
                self.prs.Close()
            if self.app:
                self.app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    def _get_save_format(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pptx":
            return 24
        if ext == ".ppt":
            return 1
        return None

    def save(self, out=None):
        path = os.path.abspath(out) if out else self.filepath
        fmt = self._get_save_format(path)
        if fmt is None:
            self.prs.SaveAs(path)
        else:
            self.prs.SaveAs(path, fmt)
        print(f"💾 已另存为: {out}" if out else "💾 已保存")

    def inspect(self):
        raw = self._run_macro("InspectPresentationJson")
        if not raw:
            return {"slides": []}
        if isinstance(raw, str):
            return json.loads(raw)
        return raw

    def print_structure(self, desc):
        for slide in desc.get("slides", []):
            print(f"\n{'=' * 50}\n📄 第 {slide['index']} 页 ({slide.get('layout', '')})\n{'=' * 50}")
            for idx, element in enumerate(slide.get("elements", []), 1):
                ph = f" [{element.get('ph_type_name','')}]" if element.get("is_placeholder") else ""
                labels = []
                if element.get("has_image"):
                    labels.append("[图片]")
                if element.get("has_chart"):
                    labels.append("[图表]")
                if element.get("has_table"):
                    labels.append("[表格]")
                if element.get("has_media"):
                    labels.append("[媒体]")
                txt = (element.get("text") or " ".join(labels) or "(无)")[:40].replace("\n", "↵")
                print(
                    f"  [{idx}] [{element.get('id')}] {element.get('name')}{ph} "
                    f"({element.get('position_label', '')}) → {txt}"
                )

    def set_notes(self, slide, text):
        return self._run_macro("SetNotes", int(slide), text)

    def append_notes(self, slide, text, separator="\n"):
        return self._run_macro("AppendNotes", int(slide), text, separator)

    def execute_action(self, action_payload):
        return self._run_macro("ExecuteActionJson", json.dumps(action_payload, ensure_ascii=False))