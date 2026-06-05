import json
import os
import subprocess
import sys

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

# 避免 emoji/中文输出在非 UTF-8 终端或 subprocess 下抛 UnicodeEncodeError
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


SUPPORTED_BACKENDS = ("pywin32", "vba", "csharp")


def add_backend_arguments(parser):
    parser.add_argument(
        "--backend",
        choices=SUPPORTED_BACKENDS,
        default=os.environ.get("PPTX_EDITOR_BACKEND", "pywin32"),
        help="底层执行策略: pywin32(默认)、vba 或 csharp",
    )
    parser.add_argument(
        "--vba-module",
        default=os.environ.get("PPTX_EDITOR_VBA_MODULE", "PptEditorBridge"),
        help="VBA backend 使用的宏模块名，默认 PptEditorBridge",
    )


def create_backend(name="pywin32", visible=False, vba_module="PptEditorBridge", csharp_host=None):
    backend = (name or "pywin32").lower()
    if backend == "pywin32":
        from pptx_editor_com import PowerPointCOM

        return PowerPointCOM(visible=visible)
    if backend == "vba":
        return PowerPointVBA(visible=visible, macro_module=vba_module)
    if backend == "csharp":
        return PowerPointCSharp(visible=visible, host_exe=csharp_host)
    raise ValueError(f"不支持的 backend: {name}")


def _resolve_csharp_host():
    """Locate the C# Interop host exe (PptInteropHost.exe).

    Search order: PPTX_EDITOR_CSHARP_HOST env var, then walk up from this file
    looking for either a bundled/published host (``csharp_host/PptInteropHost.exe``,
    used by the installed skill) or a dev build output
    (``csharp_interop/PptInteropHost/bin/<cfg>/<tfm>/PptInteropHost.exe``).
    """
    env = os.environ.get("PPTX_EDITOR_CSHARP_HOST")
    if env and os.path.isfile(env):
        return env
    cur = os.path.dirname(os.path.abspath(__file__))
    for _ in range(6):
        # Bundled / published flat layout (installed skill keeps it under csharp_host/).
        for sub in ("csharp_host", os.path.join("csharp_host", "PptInteropHost")):
            exe = os.path.join(cur, sub, "PptInteropHost.exe")
            if os.path.isfile(exe):
                return exe
        # Dev build output (repo root has csharp_interop/, skill source uses csharp_host/).
        for proj in ("csharp_interop", "csharp_host"):
            base = os.path.join(cur, proj, "PptInteropHost", "bin")
            if os.path.isdir(base):
                for cfg in ("Release", "Debug"):
                    for tfm in ("net9.0-windows", "net8.0-windows", "net10.0-windows"):
                        exe = os.path.join(base, cfg, tfm, "PptInteropHost.exe")
                        if os.path.isfile(exe):
                            return exe
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    return None


class _CSharpPrsShim:
    """Stand-in for a COM Presentation so callers that touch ``ppt.prs.Saved``
    (e.g. benchmark teardown) keep working with the C# backend."""

    Saved = False


class PowerPointCSharp:
    """PowerPoint backend driven by a persistent C# Interop host (stdio JSON-RPC).

    The C# host keeps one PowerPoint Application alive and talks late-bound COM,
    mirroring the VBA bridge protocol so benchmarks compare like-for-like."""

    def __init__(self, visible=False, host_exe=None):
        self.visible = visible
        self.prs = _CSharpPrsShim()
        self.filepath = None
        self._slide_count = 0
        self.host_exe = host_exe or _resolve_csharp_host()
        if not self.host_exe or not os.path.isfile(self.host_exe):
            raise RuntimeError(
                "找不到 C# host (PptInteropHost.exe)。请先构建:\n"
                "  dotnet build csharp_interop/PptInteropHost -c Release\n"
                "或用环境变量 PPTX_EDITOR_CSHARP_HOST 指定 exe 路径。"
            )
        self.proc = subprocess.Popen(
            [self.host_exe],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )

    def _rpc(self, payload):
        if self.proc.poll() is not None:
            stderr = self.proc.stderr.read() if self.proc.stderr else ""
            raise RuntimeError(f"C# host 进程已退出。stderr: {stderr}")
        self.proc.stdin.write(json.dumps(payload, ensure_ascii=False) + "\n")
        self.proc.stdin.flush()
        line = self.proc.stdout.readline()
        if not line:
            stderr = self.proc.stderr.read() if self.proc.stderr else ""
            raise RuntimeError(f"C# host 无响应。stderr: {stderr}")
        resp = json.loads(line)
        if not resp.get("ok"):
            raise RuntimeError(f"C# host 运行失败: {resp.get('error')}")
        return resp.get("result")

    def open(self, path):
        self.filepath = os.path.abspath(path)
        self._slide_count = self._rpc(
            {"cmd": "open", "path": self.filepath, "visible": bool(self.visible)}
        )
        print(f"📂 已打开: {self.filepath} ({self._slide_count}页)")
        return self

    def close(self):
        try:
            self._rpc({"cmd": "quit"})
        except Exception:
            pass
        try:
            self.proc.wait(timeout=10)
        except Exception:
            try:
                self.proc.terminate()
            except Exception:
                pass

    def save(self, out=None):
        path = os.path.abspath(out) if out else self.filepath
        self._rpc({"cmd": "save", "path": path})
        print(f"💾 已另存为: {out}" if out else "💾 已保存")

    def inspect(self):
        data = self._rpc({"cmd": "inspect"})
        return data or {"slides": []}

    def inspect_slide(self, slide):
        data = self._rpc({"cmd": "inspect_slide", "slide": int(slide)})
        return data or {"slides": []}

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
        return self._rpc({"cmd": "set_notes", "slide": int(slide), "text": text})

    def append_notes(self, slide, text, separator="\n"):
        return self._rpc(
            {"cmd": "append_notes", "slide": int(slide), "text": text, "separator": separator}
        )

    def execute_action(self, action_payload):
        return self._rpc({"cmd": "execute_action", "action": action_payload})


class PowerPointVBA:
    """PowerPoint backend that routes inspect/actions through Application.Run."""

    ERROR_PREFIX = "__VBA_ERROR__:"

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
            result = self.app.Run(f"{self.macro_module}.{macro_name}", *args)
            return self._normalize_macro_result(macro_name, result)
        except Exception as exc:
            raise RuntimeError(
                f"VBA backend 调用失败: {self.macro_module}.{macro_name}。"
                f"请先导入/实现 VBA 桥接模块。原始错误: {exc}"
            ) from exc

    def _normalize_macro_result(self, macro_name, result):
        if isinstance(result, str) and result.startswith(self.ERROR_PREFIX):
            raise RuntimeError(f"VBA backend 运行失败 [{macro_name}]: {result[len(self.ERROR_PREFIX):].strip()}")
        return result

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
            data = json.loads(raw)
            if isinstance(data, dict) and data.get("error"):
                raise RuntimeError(f"VBA backend inspect 失败: {data['error']}")
            return data
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