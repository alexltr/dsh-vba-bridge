"""Probe: offline VBA extraction with oletools (no Excel)."""
import sys, json, zipfile, io
sys.path.insert(0, r"F:\MyProjects\dsh\vba_bridge\vendor")
import olefile
from oletools.olevba import VBA_Parser

path = sys.argv[1]
print("== container ==")
if zipfile.is_zipfile(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        print("zip entries:", len(names))
        for n in names:
            if "vba" in n.lower():
                print("  ", n, z.getinfo(n).file_size)
        blob = z.read("xl/vbaProject.bin")
    print("== OLE streams inside xl/vbaProject.bin ==")
    ole = olefile.OleFileIO(io.BytesIO(blob))
    for entry in ole.listdir():
        print("  " + "/".join(entry), ole.get_size(entry))
    ole.close()
else:
    print("OLE (legacy .xls/.xlsb) container")

print("== olevba ==")
p = VBA_Parser(path)
print("container type:", p.type)
mods = []
for (filename, stream, vba_filename, code) in p.extract_macros():
    mods.append({"stream": "/".join(stream), "module": vba_filename, "chars": len(code)})
    print("-" * 60)
    print("MODULE:", vba_filename, "| stream:", "/".join(stream), "| chars:", len(code))
    print(code[:400])
p.close()
print("== summary ==")
print(json.dumps(mods, ensure_ascii=False, indent=2))
