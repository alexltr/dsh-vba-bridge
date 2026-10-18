"""PoC: prove the edited VBA compiles and runs inside Excel."""
import json, os, shutil, time, traceback
import win32com.client as win32

SRC = r"F:\MyProjects\dsh\vba_bridge\work\tool-write.xlsm"
DST = r"F:\MyProjects\dsh\vba_bridge\work\run-test.xlsm"
out = {}
shutil.copy2(SRC, DST)

app = win32.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
app.EnableEvents = False
app.AutomationSecurity = 1  # msoAutomationSecurityLow: allow the macro we just injected
wb = None
try:
    wb = app.Workbooks.Open(DST, UpdateLinks=0, ReadOnly=False)
    result = app.Run("ToolSmoke", 21)
    out["ToolSmoke(21)"] = result
    out["expected"] = 42
    out["compiled_and_ran"] = bool(result == 42)
    # also confirm the untouched original procedure still exists and compiles
    app.Run("DshPong")
    out["original_macro_ran"] = True
    wb.Close(SaveChanges=False)
    wb = None
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
