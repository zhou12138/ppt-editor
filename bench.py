"""
Benchmark: pywin32 vs VBA vs C# Interop (exe) vs C# add-in vs Python add-in.

Five backends arranged as a (Python / C#) x (out-of-process / in-process) matrix
(VBA is the in-process reference point):
  - OUT-of-process (cross-process COM marshalling per op):
        pywin32       -> python.exe  driving PowerPoint over COM
        csharp        -> C# exe host driving PowerPoint over COM
  - IN-process (zero cross-process marshalling; logic runs inside POWERPNT.EXE):
        vba           -> VBA macro bridge
        csharp-addin  -> C# COM add-in bridge (early-bound interop)
        pywin32-addin -> Python (pywin32) COM add-in bridge

The matrix isolates the variable: same language, in-proc vs out-proc, shows the
cross-process IPC — not the language — is what made pywin32 slow.

Run: python bench.py pywin32 | vba | csharp | csharp-addin | pywin32-addin
"""
import sys, os, time, json

sys.path.insert(0, r"C:\edge_workspace_1\ppt-editor\skills\pptx-local-offline\scripts")
from ppt_backend import create_backend

PPTX = r"C:\Users\zhouweiwei.REDMOND\OneDrive - Microsoft\Desktop\柿子成长过程概述.pptx"

TITLE_SHAPES = {1:"TextBox 2", 2:"TextBox 6", 3:"TextBox 6", 4:"TextBox 6",
                5:"TextBox 6", 6:"TextBox 6", 7:"TextBox 6", 8:"TextBox 11"}

def fmt(sec): return f"{sec:.3f}s"

def run(backend_name):
    results = {}
    ppt = create_backend(backend_name, visible=True)

    # open
    t0 = time.perf_counter()
    ppt.open(PPTX)
    results["open"] = time.perf_counter() - t0
    print(f"  open                      {fmt(results['open'])}")

    try:
        # inspect
        t0 = time.perf_counter()
        desc = ppt.inspect()
        results["inspect"] = time.perf_counter() - t0
        sc = len(desc.get("slides",[])); shapes = sum(len(s.get("elements",[])) for s in desc.get("slides",[]))
        print(f"  inspect                   {fmt(results['inspect'])}  ({sc} slides, {shapes} shapes)")

        if backend_name == "pywin32":
            # find 8 titles
            t0 = time.perf_counter()
            found = []
            for si, nm in TITLE_SHAPES.items():
                found.extend(ppt.find_shape(si, {"name": nm}) or [])
            results["find_8_titles"] = time.perf_counter() - t0
            print(f"  find 8 titles             {fmt(results['find_8_titles'])}  ({len(found)} found)")

            # modify_font single
            s1 = ppt.find_shape(1, {"name": "TextBox 2"})
            if s1:
                t0 = time.perf_counter()
                ppt.modify_font(s1[0], bold=True, color=0x0000FF)
                results["modify_font_1"] = time.perf_counter() - t0
                print(f"  modify_font single        {fmt(results['modify_font_1'])}")

            # modify_font 8 pages
            t0 = time.perf_counter()
            for si, nm in TITLE_SHAPES.items():
                ss = ppt.find_shape(si, {"name": nm})
                for s in (ss or []):
                    ppt.modify_font(s, bold=True)
            results["modify_font_8"] = time.perf_counter() - t0
            print(f"  modify_font 8 pages       {fmt(results['modify_font_8'])}")

            # modify_text
            s4 = ppt.find_shape(1, {"name": "TextBox 4"})
            if s4:
                t0 = time.perf_counter()
                ppt.modify_text(s4[0], "BenchmarkText")
                results["modify_text"] = time.perf_counter() - t0
                print(f"  modify_text               {fmt(results['modify_text'])}")
                ppt.modify_text(s4[0], "农业科普教育")

            # move_shape
            if s4:
                t0 = time.perf_counter()
                ppt.move_shape(s4[0], left=450, top=482)
                results["move_shape"] = time.perf_counter() - t0
                print(f"  move_shape                {fmt(results['move_shape'])}")

            # resize_shape
            if s4:
                t0 = time.perf_counter()
                ppt.resize_shape(s4[0], width=78, height=15)
                results["resize_shape"] = time.perf_counter() - t0
                print(f"  resize_shape              {fmt(results['resize_shape'])}")

            # add + delete textbox
            t0 = time.perf_counter()
            ppt.add_textbox(1, "BM_TEST", 100, 100, 200, 40)
            ff = ppt.find_shape(1, {"text": "BM_TEST"})
            if ff: ppt.delete_shape(ff[0])
            results["add_del_textbox"] = time.perf_counter() - t0
            print(f"  add+delete textbox        {fmt(results['add_del_textbox'])}")

            # set_notes
            t0 = time.perf_counter()
            ppt.set_notes(1, "benchmark note")
            results["set_notes"] = time.perf_counter() - t0
            print(f"  set_notes                 {fmt(results['set_notes'])}")

            # batch mixed
            t0 = time.perf_counter()
            for si, nm in TITLE_SHAPES.items():
                ss = ppt.find_shape(si, {"name": nm})
                if ss:
                    try: ppt.modify_font(ss[0], bold=True)
                    except: pass
            ppt.set_notes(1, "batch done")
            results["batch_mixed"] = time.perf_counter() - t0
            print(f"  batch mixed               {fmt(results['batch_mixed'])}")

        else:  # vba / csharp / csharp-codeact / *-addin — drive ops through execute_action
                # (csharp-codeact transparently routes each action through execute_code/Roslyn)
            # find via inspect (already done)
            results["find_8_titles"] = results["inspect"]
            print(f"  find (via inspect)        {fmt(results['find_8_titles'])}")

            # modify_font single
            t0 = time.perf_counter()
            ppt.execute_action({"action":"modify_font","slide":1,"target":{"name":"TextBox 2"},"params":{"bold":True,"color":255}})
            results["modify_font_1"] = time.perf_counter() - t0
            print(f"  modify_font single        {fmt(results['modify_font_1'])}")

            # modify_font 8 pages
            t0 = time.perf_counter()
            for si, nm in TITLE_SHAPES.items():
                ppt.execute_action({"action":"modify_font","slide":si,"target":{"name":nm},"params":{"bold":True}})
            results["modify_font_8"] = time.perf_counter() - t0
            print(f"  modify_font 8 pages       {fmt(results['modify_font_8'])}")

            # modify_text
            t0 = time.perf_counter()
            ppt.execute_action({"action":"modify_text","slide":1,"target":{"name":"TextBox 4"},"params":{"text":"BenchmarkText"}})
            results["modify_text"] = time.perf_counter() - t0
            print(f"  modify_text               {fmt(results['modify_text'])}")
            ppt.execute_action({"action":"modify_text","slide":1,"target":{"name":"TextBox 4"},"params":{"text":"农业科普教育"}})

            # move_shape
            t0 = time.perf_counter()
            ppt.execute_action({"action":"move_shape","slide":1,"target":{"name":"TextBox 4"},"params":{"left":450,"top":482}})
            results["move_shape"] = time.perf_counter() - t0
            print(f"  move_shape                {fmt(results['move_shape'])}")

            # resize_shape
            t0 = time.perf_counter()
            ppt.execute_action({"action":"resize_shape","slide":1,"target":{"name":"TextBox 4"},"params":{"width":78,"height":15}})
            results["resize_shape"] = time.perf_counter() - t0
            print(f"  resize_shape              {fmt(results['resize_shape'])}")

            # add + delete textbox
            t0 = time.perf_counter()
            ppt.execute_action({"action":"add_textbox","slide":1,"params":{"text":"BM_TEST","left":100,"top":100,"width":200,"height":40}})
            ppt.execute_action({"action":"delete_shape","slide":1,"target":{"text":"BM_TEST"}})
            results["add_del_textbox"] = time.perf_counter() - t0
            print(f"  add+delete textbox        {fmt(results['add_del_textbox'])}")

            # set_notes
            t0 = time.perf_counter()
            ppt.set_notes(1, "benchmark note")
            results["set_notes"] = time.perf_counter() - t0
            print(f"  set_notes                 {fmt(results['set_notes'])}")

            # batch mixed
            t0 = time.perf_counter()
            for si, nm in TITLE_SHAPES.items():
                try:
                    ppt.execute_action({"action":"modify_font","slide":si,"target":{"name":nm},"params":{"bold":True}})
                except Exception:
                    pass  # some shapes may not match after add/delete
            ppt.set_notes(1, "batch done")
            results["batch_mixed"] = time.perf_counter() - t0
            print(f"  batch mixed               {fmt(results['batch_mixed'])}")

            if backend_name == "csharp-codeact" and hasattr(ppt, "run_actions"):
                # CodeAct's win: the same 8 modify_font ops compiled into ONE C#
                # script and run in a SINGLE round-trip (vs modify_font_8 above,
                # which is 8 separate execute_code round-trips).
                t0 = time.perf_counter()
                ppt.run_actions([
                    {"action": "modify_font", "slide": si, "target": {"name": nm}, "params": {"bold": True}}
                    for si, nm in TITLE_SHAPES.items()
                ])
                results["batch8_codeact_1rt"] = time.perf_counter() - t0
                print(f"  batch 8 (codeact 1-rt)    {fmt(results['batch8_codeact_1rt'])}")

    finally:
        try:
            ppt.prs.Saved = True
            ppt.close()
        except: pass

    total = sum(v for k,v in results.items() if k != "open")
    results["total"] = total
    print(f"  {'─'*40}")
    print(f"  TOTAL (excl open)         {fmt(total)}")
    return results


if __name__ == "__main__":
    backend = sys.argv[1] if len(sys.argv) > 1 else "pywin32"
    print(f"{'='*60}")
    print(f"  Benchmark: {backend.upper()}")
    print(f"  File: {os.path.basename(PPTX)} (8 slides, 119 shapes)")
    print(f"{'='*60}")
    r = run(backend)
    # Write results to JSON for comparison
    out = f"bench_{backend}.json"
    with open(os.path.join(r"C:\edge_workspace_1\ppt-editor", out), "w") as f:
        json.dump(r, f, indent=2)
    print(f"\n  Results saved to {out}")
