"""PoC: read + edit VBA source through Excel COM, then verify the saved file."""
import hashlib, json, os, shutil, sys, time, traceback
import win32com.client as win32

SRC = r"F:\MyProjects\dsh\vba_bridge\fixtures\sample.xlsm"
DST = r"F:\MyProjects\dsh\vba_bridge\work\edit1.xlsm"
MODULE = "模块1"

NEW_PROC = '''
Sub DshProbe()
    ' injected by dsh vba_bridge PoC
    Debug.Print "hello from dsh"
End Sub
'''

def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()

os.makedirs(os.path.dirname(DST), exist_ok=True)
shutil.copy2(SRC, DST)
out = {"src": SRC, "dst": DST, "sha256_before": sha(SRC)}

t0 = time.time()
app = win32.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
app.AutomationSecurity = 3
app.EnableEvents = False
app.AskToUpdateLinks = False
wb = None
try:
    wb = app.Workbooks.Open(DST, UpdateLinks=0, ReadOnly=False)
    out["opened_in_s"] = round(time.time() - t0, 1)
    vbp = wb.VBProject
    cm = vbp.VBComponents(MODULE).CodeModule
    before = cm.Lines(1, cm.CountOfLines)
    out["module"] = MODULE
    out["lines_before"] = int(cm.CountOfLines)
    out["code_before_tail"] = before[-200:]

    # idempotent append: only add if marker absent
    if "DshProbe" not in before:
        cm.AddFromString(NEW_PROC)
    out["lines_after"] = int(cm.CountOfLines)
    out["code_after_tail"] = cm.Lines(1, cm.CountOfLines)[-400:]

    wb.Save()
    out["saved_in_s"] = round(time.time() - t0, 1)
    out["sha256_after"] = sha(DST)
    wb.Close(SaveChanges=False)
    wb = None
except Exception as exc:
    out["error"] = repr(exc)
    out["trace"] = traceback.format_exc()
finally:
    if wb is not None:
        try: wb.Close(SaveChanges=False)
        except Exception: pass
    try: app.Quit()
    except Exception: pass
    time.sleep(0.5)

# reopen and verify through COM
app2 = win32.DispatchEx("Excel.Application")
app2.Visible = False
app2.DisplayAlerts = False
app2.AutomationSecurity = 3
app2.EnableEvents = False
wb2 = None
try:
    wb2 = app2.Workbooks.Open(DST, UpdateLinks=0, ReadOnly=True)
    cm2 = wb2.VBProject.VBComponents(MODULE).CodeModule
    code2 = cm2.Lines(1, cm2.CountOfLines)
    out["reopened_has_probe"] = "DshProbe" in code2
    out["reopened_lines"] = int(cm2.CountOfLines)
    out["reopened_tail"] = code2[-400:]
    wb2.Close(SaveChanges=False)
    wb2 = None
    out["ok"] = "error" not in out
except Exception as exc:
    out["reopen_error"] = repr(exc)
    out["ok"] = False
finally:
    if wb2 is not None:
        try: wb2.Close(SaveChanges=False)
        except Exception: pass
    try: app2.Quit()
    except Exception: pass

print(json.dumps(out, ensure_ascii=False, indent=2))
