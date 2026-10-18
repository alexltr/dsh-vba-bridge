"""Probe: UserForm code-behind + VBComponent.Export behaviour."""
import json, os, shutil, tempfile, traceback
import win32com.client as win32

SRC = r"F:\MyProjects\dsh\vba_bridge\test_samples\MyVBALibrary.xlam"
WORK = r"F:\MyProjects\dsh\vba_bridge\work\form-probe.xlam"
OUT = r"F:\MyProjects\dsh\vba_bridge\work\form-export"
shutil.copy2(SRC, WORK)
shutil.rmtree(OUT, ignore_errors=True)
os.makedirs(OUT, exist_ok=True)

out = {}
app = win32.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
app.AutomationSecurity = 3
app.EnableEvents = False
wb = None
try:
    wb = app.Workbooks.Open(WORK, UpdateLinks=0, ReadOnly=False, Notify=False)
    vbp = wb.VBProject
    form = vbp.VBComponents("FRM_PictureViewer")
    cm = form.CodeModule
    out["form"] = {
        "name": form.Name,
        "type": int(form.Type),
        "lines": int(cm.CountOfLines),
        "has_designer": bool(form.Designer),
        "designer": str(form.Designer) if form.Designer else None,
    }
    code = cm.Lines(1, int(cm.CountOfLines)) if cm.CountOfLines else ""
    out["form"]["code_head"] = code[:600]
    out["form"]["has_click_handler"] = "_Click" in code

    # write a harmless comment into the form's code (on a copy) and save
    cm.AddFromString("\r\n' form code touched by bridge probe\r\n")
    wb.Save()
    out["form"]["write_ok"] = "form code touched" in cm.Lines(1, cm.CountOfLines)

    # VBComponent.Export for each component type
    exported = []
    for comp in vbp.VBComponents:
        ext = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}.get(int(comp.Type), ".txt")
        target = os.path.join(OUT, comp.Name + ext)
        try:
            comp.Export(target)
            extra = sorted(
                f for f in os.listdir(OUT)
                if f.startswith(comp.Name) and f != os.path.basename(target)
            )
            exported.append({
                "module": comp.Name, "type": int(comp.Type), "target": os.path.basename(target),
                "exists": os.path.exists(target),
                "bytes": os.path.getsize(target) if os.path.exists(target) else 0,
                "sidecars": extra,
            })
        except Exception as exc:
            exported.append({"module": comp.Name, "error": repr(exc)})
    out["export"] = exported
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

print(json.dumps(out, ensure_ascii=False, indent=2))
