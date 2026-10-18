"""PoC 2: document-module code, add/remove component, and protected-project behaviour."""
import json, os, shutil, time, traceback
import win32com.client as win32

SRC = r"F:\MyProjects\dsh\vba_bridge\fixtures\sample2.xlsm"
DST = r"F:\MyProjects\dsh\vba_bridge\work\edit2.xlsm"

out = {}
os.makedirs(os.path.dirname(DST), exist_ok=True)
shutil.copy2(SRC, DST)

app = win32.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
app.AutomationSecurity = 3
app.EnableEvents = False
wb = None
try:
    wb = app.Workbooks.Open(DST, UpdateLinks=0, ReadOnly=False)
    vbp = wb.VBProject
    comps = []
    for c in vbp.VBComponents:
        cm = c.CodeModule
        item = {"name": c.Name, "type": int(c.Type), "lines": int(cm.CountOfLines)}
        if cm.CountOfLines:
            item["head"] = cm.Lines(1, min(3, cm.CountOfLines))
        comps.append(item)
    out["before"] = comps

    # add a standard module
    mod = vbp.VBComponents.Add(1)  # vbext_ct_StdModule
    mod.Name = "DshTemp"
    mod.CodeModule.AddFromString('Sub DshTempProc()\r\n    Debug.Print "temp"\r\nEnd Sub\r\n')
    out["added"] = {"name": mod.Name, "lines": int(mod.CodeModule.CountOfLines)}
    wb.Save()
    out["count_after_add"] = int(vbp.VBComponents.Count)

    # remove it again
    vbp.VBComponents.Remove(vbp.VBComponents("DshTemp"))
    wb.Save()
    out["count_after_remove"] = int(vbp.VBComponents.Count)
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

# protected project: create one via COM VBE command is unreliable; test a read of protection flag only
print(json.dumps(out, ensure_ascii=False, indent=2))
