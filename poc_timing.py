"""Where does bridge latency go?"""
import json, time, win32com.client as win32

path = r"F:\MyProjects\dsh\vba_bridge\fixtures\sample.xlsm"
t = {}
t0 = time.time()
app = win32.DispatchEx("Excel.Application")
t["dispatch"] = round(time.time() - t0, 2)
app.Visible = False
app.DisplayAlerts = False
app.AutomationSecurity = 3
app.EnableEvents = False
t1 = time.time()
wb = app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True, Notify=False)
t["open"] = round(time.time() - t1, 2)
t1 = time.time()
vbp = wb.VBProject
n = vbp.VBComponents.Count
t["project"] = round(time.time() - t1, 2)
t1 = time.time()
codes = [(c.Name, c.CodeModule.CountOfLines) for c in vbp.VBComponents]
t["components"] = round(time.time() - t1, 2)
t1 = time.time()
wb.Close(SaveChanges=False)
t["close"] = round(time.time() - t1, 2)
t1 = time.time()
app.Quit()
t["quit"] = round(time.time() - t1, 2)
t["total"] = round(time.time() - t0, 2)
t["components_count"] = n
print(json.dumps(t, indent=2))
