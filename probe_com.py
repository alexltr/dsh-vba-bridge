"""Probe: open an .xlsm through Excel COM and enumerate its VBA project."""
import sys, json, traceback
import win32com.client as win32

path = sys.argv[1]

app = win32.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
app.AutomationSecurity = 3          # msoAutomationSecurityForceDisable
app.EnableEvents = False
app.AskToUpdateLinks = False

out = {"path": path, "excel": app.Version}
wb = None
try:
    wb = app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
    vbp = wb.VBProject
    out["project"] = vbp.Name
    try:
        out["protection"] = int(vbp.Protection)
    except Exception as exc:
        out["protection"] = f"error: {exc}"
    comps = []
    for c in vbp.VBComponents:
        entry = {"name": c.Name, "type": int(c.Type), "lines": int(c.CodeModule.CountOfLines)}
        try:
            entry["designer"] = bool(c.Designer)
        except Exception as exc:
            entry["designer"] = f"error: {exc}"
        comps.append(entry)
    out["components"] = comps
    out["references"] = [
        {"name": r.Name, "version": r.Major, "minor": r.Minor, "builtin": bool(r.BuiltIn)}
        for r in vbp.References
    ]
    out["ok"] = True
except Exception as exc:
    out["ok"] = False
    out["error"] = repr(exc)
    out["trace"] = traceback.format_exc()
finally:
    if wb is not None:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
    try:
        app.Quit()
    except Exception:
        pass
print(json.dumps(out, ensure_ascii=False, indent=2))
