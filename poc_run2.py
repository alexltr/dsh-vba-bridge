"""PoC: prove tool-written VBA compiles and runs inside Excel."""
import json, shutil, time, traceback
import win32com.client as win32

SRC = r"F:\MyProjects\dsh\vba_bridge\work\run-test.xlsm"
DST = r"F:\MyProjects\dsh\vba_bridge\work\run-test2.xlsm"
out = {}
shutil.copy2(SRC, DST)

app = win32.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
app.EnableEvents = False
app.AutomationSecurity = 1  # allow the macro written by the bridge
wb = None
try:
    wb = app.Workbooks.Open(DST, UpdateLinks=0, ReadOnly=False)
    out["modules"] = [str(c.Name) for c in wb.VBProject.VBComponents if int(c.CodeModule.CountOfLines) > 0]
    out["DshDouble(21)"] = app.Run("DshDouble", 21)
    out["expected"] = 42
    out["compiled_and_ran"] = bool(out["DshDouble(21)"] == 42)
    out["ok"] = True
except Exception as exc:
    out["error"] = repr(exc)
    out["trace"] = traceback.format_exc()
    out["ok"] = False
finally:
    if wb is not None:
        try: wb.Close(SaveChanges=False)
        except Exception: pass
    try: app.Quit()
    except Exception: pass
    time.sleep(0.3)

print(json.dumps(out, ensure_ascii=False, indent=2))
