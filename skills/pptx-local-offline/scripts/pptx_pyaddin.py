"""In-process PowerPoint COM add-in implemented in PURE PYTHON (pywin32).

This is the 5th backend strategy and the in-process twin of the pywin32
(out-of-process) backend. It runs the editing logic INSIDE POWERPNT.EXE via a
COM add-in — exactly like VBA and the C# add-in — so the PowerPoint object model
is touched with ZERO cross-process marshalling.

Why it matters (the benchmark thesis): the only thing that made pywin32 slow was
the out-of-process COM IPC (one round-trip per property/method). Move the SAME
Python code in-process and the per-shape traversal becomes free, proving the
bottleneck is the IPC boundary, not the language. This fills the Python x
in-process cell of the 2x2 (Python / C#) x (in-proc / out-proc) matrix.

How external drivers reach it (mirrors the C# add-in):
    app.COMAddIns.Item("PptEditor.PyAddIn").Object  -> the bridge
Then ONE coarse cross-process call per operation:
    bridge.Ping()  bridge.InspectJson()  bridge.InspectSlideJson(i)
    bridge.ExecuteActionJson(json)
All per-shape work happens in-process by reusing the proven PowerPointCOM +
_dispatch logic, bound to the in-process Application.

Implementation notes:
- IDTExtensibility2 is provided by win32com.universal against the Add-In Designer
  type library {AC0714F2-3D04-11D1-AE7D-00A0C90F26F4} (ships with Office), so the
  Office lifecycle callbacks (OnConnection ...) marshal correctly.
- EventHandlerPolicy maps the incoming IDTExtensibility2 vtable/dispatch calls to
  the matching Python methods; _public_methods_ exposes the automation bridge.
- PowerPoint add-ins only load under an INTERACTIVE launch (automation / DispatchEx
  refuses to connect them), identical to the C# add-in. Headless CI won't load it.

Register / unregister (per-user, no admin):
    python pptx_pyaddin.py                # register COM server + PowerPoint add-in
    python pptx_pyaddin.py --unregister   # remove both
"""

import json
import os
import sys

import pythoncom
import win32com.server.util
from win32com import universal

# IDTExtensibility2 lives in the Microsoft Add-In Designer type library. Building
# it here lets pywin32 marshal the Office add-in lifecycle interface by IID.
universal.RegisterInterfaces(
    "{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}", 0, 1, 0, ["_IDTExtensibility2"]
)

# New, dedicated identity (distinct from the C# add-in's CLSID/ProgId).
PYADDIN_CLSID = "{5F3D2E1A-9C4B-4A77-B2E6-7D1F8E0A6C42}"
PYADDIN_PROGID = "PptEditor.PyAddIn"

_LOG_PATH = os.path.join(
    os.environ.get("TEMP", os.path.expanduser("~")), "ppteditor_pyaddin.log"
)


def _log(msg):
    try:
        import datetime

        with open(_LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(datetime.datetime.now().strftime("%H:%M:%S.%f ") + str(msg) + "\n")
    except Exception:
        pass


def _import_engine():
    """Import the reusable editing logic from sibling modules.

    When PowerPoint loads this add-in in-process, the registered ``PythonPath``
    (the directory of this file) is on sys.path, so its siblings import too.
    """
    here = os.path.dirname(os.path.abspath(__file__))
    if here not in sys.path:
        sys.path.insert(0, here)
    from pptx_editor_com import PowerPointCOM
    from pptx_editor_llm import _dispatch

    return PowerPointCOM, _dispatch


class PptEditorPyAddin:
    """The in-process add-in + automation bridge (one object, two roles)."""

    # --- COM registration metadata (consumed by win32com.server.register) ---
    _reg_clsid_ = PYADDIN_CLSID
    _reg_progid_ = PYADDIN_PROGID
    _reg_desc_ = "PptEditor Python in-process add-in"
    # Explicit spec so the in-process host re-imports the module (not __main__).
    _reg_class_spec_ = "pptx_pyaddin.PptEditorPyAddin"
    _reg_policy_spec_ = "win32com.server.policy.EventHandlerPolicy"

    # IDTExtensibility2 callbacks are dispatched to the like-named methods.
    _com_interfaces_ = ["_IDTExtensibility2"]
    # Automation surface exposed via COMAddIn.Object for external drivers.
    _public_methods_ = [
        "Ping",
        "InspectJson",
        "InspectSlideJson",
        "ExecuteActionJson",
        "RequestComAddInAutomationService",
    ]

    def __init__(self):
        self._app = None
        self._ppt = None  # PowerPointCOM bound to the in-process Application
        self._dispatch = None

    # ----------------------------- IDTExtensibility2 -----------------------------

    def OnConnection(self, application, connectMode, addin, custom):
        _log("OnConnection connectMode=%s" % connectMode)
        self._app = application
        try:
            PowerPointCOM, dispatch = _import_engine()
            self._dispatch = dispatch
            # Bind the proven engine to the in-process app WITHOUT spawning a
            # second PowerPoint (bypass PowerPointCOM.__init__'s Dispatch()).
            ppt = PowerPointCOM.__new__(PowerPointCOM)
            ppt.app = application
            ppt.prs = None
            ppt.filepath = None
            self._ppt = ppt
            _log("engine bound OK")
        except Exception as exc:
            _log("engine bind FAILED: %s" % exc)
        try:
            # Publish the automation bridge so external drivers can reach it via
            # Application.COMAddIns.Item(ProgId).Object.
            addin.Object = win32com.server.util.wrap(self)
            _log("addin.Object set OK")
        except Exception as exc:
            _log("addin.Object set FAILED: %s" % exc)

    def OnDisconnection(self, mode, custom):
        _log("OnDisconnection mode=%s" % mode)
        self._ppt = None
        self._app = None

    def OnAddInsUpdate(self, custom):
        pass

    def OnStartupComplete(self, custom):
        pass

    def OnBeginShutdown(self, custom):
        pass

    def RequestComAddInAutomationService(self):
        return win32com.server.util.wrap(self)

    # ----------------------------- automation bridge -----------------------------

    def _bind_active(self):
        """Point the engine at the active (or first) presentation for this call."""
        if self._ppt is None:
            raise RuntimeError("engine not initialized (OnConnection did not run)")
        prs = None
        try:
            prs = self._app.ActivePresentation
        except Exception:
            prs = None
        if prs is None:
            prs = self._app.Presentations(1)
        self._ppt.prs = prs
        try:
            self._ppt.filepath = prs.FullName
        except Exception:
            self._ppt.filepath = None
        return self._ppt

    def Ping(self):
        return "pong"

    def InspectJson(self):
        ppt = self._bind_active()
        return json.dumps(ppt.inspect(), ensure_ascii=False)

    def InspectSlideJson(self, slideIndex):
        ppt = self._bind_active()
        full = ppt.inspect()
        idx = int(slideIndex)
        slides = [s for s in full.get("slides", []) if s.get("index") == idx]
        return json.dumps({"slides": slides}, ensure_ascii=False)

    def ExecuteActionJson(self, actionJson):
        ppt = self._bind_active()
        action = json.loads(actionJson)
        # Accept both nested {"action": {...}} and flat {"action": "name", ...}.
        if isinstance(action.get("action"), dict):
            action = action["action"]
        name = action.get("action", "")
        slide = action.get("slide")
        target = action.get("target", {})
        params = action.get("params", {})

        # Notes are not part of _dispatch; route them to PowerPointCOM directly.
        if name == "set_notes":
            result = ppt.set_notes(int(slide or 1), params.get("text", ""))
        elif name == "append_notes":
            result = ppt.append_notes(
                int(slide or 1), params.get("text", ""), params.get("separator", "\n")
            )
        else:
            result = self._dispatch(ppt, name, slide, target, params)
        return result if isinstance(result, str) else json.dumps(result, ensure_ascii=False)


# --------------------------- register / unregister ---------------------------

_ADDINS_KEY = r"Software\Microsoft\Office\PowerPoint\Addins"


def _register_addin(klass):
    import winreg

    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _ADDINS_KEY + "\\" + klass._reg_progid_) as sub:
        winreg.SetValueEx(sub, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
        winreg.SetValueEx(sub, "LoadBehavior", 0, winreg.REG_DWORD, 3)
        winreg.SetValueEx(sub, "Description", 0, winreg.REG_SZ, klass._reg_desc_)
        winreg.SetValueEx(sub, "FriendlyName", 0, winreg.REG_SZ, "PptEditor Python Add-In")
    print(
        "Registered '%s' (CLSID %s) per-user. LoadBehavior=3."
        % (klass._reg_progid_, klass._reg_clsid_)
    )


def _unregister_addin(klass):
    import winreg

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, _ADDINS_KEY + "\\" + klass._reg_progid_)
        print("Removed PowerPoint add-in key '%s'." % klass._reg_progid_)
    except OSError:
        pass


if __name__ == "__main__":
    import win32com.server.register

    # Registers (default) or unregisters (--unregister / --clean) the COM server.
    win32com.server.register.UseCommandLine(PptEditorPyAddin)
    if "--unregister" in sys.argv or "--clean" in sys.argv:
        _unregister_addin(PptEditorPyAddin)
    else:
        _register_addin(PptEditorPyAddin)
