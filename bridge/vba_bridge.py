"""vba_bridge - read and edit Excel VBA projects.

Primary engine: Excel COM automation (Windows, Excel installed, Trust Center
"Trust access to the VBA project object model" enabled).  Fallback engine:
offline parsing of the OLE/OVBA streams with oletools (read-only).

JSON protocol (used by the DSH plugin):
    python vba_bridge.py --json-stdin   < request.json   > response.json

CLI (for humans):
    python vba_bridge.py inspect FILE [--engine auto|com|ovba]
    python vba_bridge.py read    FILE MODULE [--engine auto|com|ovba]
    python vba_bridge.py write   FILE MODULE --code-file PATH [--mode replace|append] [--create]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import re
import shutil
import sys
import tempfile
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
VENDOR = os.path.join(os.path.dirname(HERE), "vendor")

TYPE_NAMES = {1: "std", 2: "class", 3: "form", 100: "document"}
TYPE_EXT = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}

EXCEL_EXTENSIONS = {".xls", ".xlsm", ".xlsb", ".xlam", ".xltm", ".xltx", ".xlsm"}


class BridgeError(Exception):
    def __init__(self, code: str, message: str, hint: str | None = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.hint = hint


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def norm_code(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def to_crlf(text: str) -> str:
    return norm_code(text).replace("\n", "\r\n")


def ensure_trailing_newline(text: str) -> str:
    text = to_crlf(text)
    return text if text.endswith("\r\n") else text + "\r\n"


def validate_file(path: str) -> str:
    if not path:
        raise BridgeError("INVALID_REQUEST", "file is required")
    path = os.path.abspath(os.path.expandvars(os.path.expanduser(path)))
    if not os.path.isfile(path):
        raise BridgeError("FILE_NOT_FOUND", f"file not found: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext in {".xlsx", ".xltx"}:
        raise BridgeError(
            "NOT_MACRO_ENABLED",
            f"{os.path.basename(path)} is an OOXML workbook without a VBA project container",
            "use an .xlsm/.xlsb/.xls file; .xlsx cannot store VBA",
        )
    return path


# --------------------------------------------------------------------------- COM


def _com_error_text(exc: BaseException) -> str:
    parts = [str(exc)]
    info = getattr(exc, "excepinfo", None)
    if info:
        try:
            if info[2]:
                parts.append(str(info[2]))
        except Exception:
            pass
    return " | ".join(dict.fromkeys(p for p in parts if p))


def _map_com_error(exc: BaseException) -> BridgeError:
    text = _com_error_text(exc)
    low = text.lower()
    if "trust access" in low or "programmatic access" in low:
        return BridgeError(
            "VBA_TRUST_DISABLED",
            "Excel blocked programmatic access to the VBA project",
            'enable Excel > File > Options > Trust Center > Trust Center Settings > '
            'Macro Settings > "Trust access to the VBA project object model" '
            "(HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Security\\AccessVBOM=1)",
        )
    if "project is protected" in low or "unviewable" in low or "password" in low:
        return BridgeError(
            "PROJECT_PROTECTED",
            "the VBA project is password-protected; editing requires removing the password in the VBE",
        )
    if "being used" in low or "in use" in low or "locked" in low or "read-only" in low:
        return BridgeError("FILE_LOCKED", "the workbook is open or locked by another process")
    return BridgeError("COM_ERROR", text)


class Excel:
    """One hidden Excel instance, closed deterministically."""

    def __init__(self):
        try:
            import win32com.client  # noqa: F401
        except Exception as exc:  # pragma: no cover - environment dependent
            raise BridgeError(
                "PYWIN32_MISSING",
                "pywin32 is not installed for this Python: " + str(exc),
                "pip install pywin32",
            ) from exc
        import win32com.client as win32

        try:
            self.app = win32.DispatchEx("Excel.Application")
        except Exception as exc:
            raise BridgeError(
                "EXCEL_UNAVAILABLE",
                "cannot start Excel via COM: " + _com_error_text(exc),
                "install Microsoft Excel (desktop) and make sure it is not blocked by policy",
            ) from exc
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.EnableEvents = False
        self.app.AskToUpdateLinks = False
        try:
            # msoAutomationSecurityForceDisable: macros never run during a bridge call.
            self.app.AutomationSecurity = 3
        except Exception:
            pass
        self.wb = None

    def open(self, path: str, readonly: bool):
        try:
            self.wb = self.app.Workbooks.Open(
                path, UpdateLinks=0, ReadOnly=readonly, Notify=False
            )
        except Exception as exc:
            raise _map_com_error(exc) from exc
        return self.wb

    def project(self):
        try:
            return self.wb.VBProject
        except Exception as exc:
            raise _map_com_error(exc) from exc

    def close(self):
        if self.wb is not None:
            try:
                self.wb.Close(SaveChanges=False)
            except Exception:
                pass
            self.wb = None
        try:
            self.app.Quit()
        except Exception:
            pass


def _component_type(component) -> str:
    try:
        return TYPE_NAMES.get(int(component.Type), f"type{int(component.Type)}")
    except Exception:
        return "unknown"


def com_inspect(path: str) -> dict:
    xl = Excel()
    try:
        xl.open(path, readonly=True)
        vbp = xl.project()
        components = []
        try:
            comps = list(vbp.VBComponents)
        except Exception as exc:
            raise _map_com_error(exc) from exc
        for c in comps:
            try:
                lines = int(c.CodeModule.CountOfLines)
            except Exception:
                lines = -1
            components.append(
                {
                    "name": str(c.Name),
                    "type": _component_type(c),
                    "type_id": int(c.Type),
                    "lines": lines,
                }
            )
        references = []
        try:
            for r in vbp.References:
                references.append(
                    {
                        "name": str(r.Name),
                        "version": f"{int(r.Major)}.{int(r.Minor)}",
                        "builtin": bool(r.BuiltIn),
                        "guid": str(getattr(r, "GUID", "") or ""),
                    }
                )
        except Exception:
            pass
        protection = 0
        try:
            protection = int(vbp.Protection)
        except Exception:
            pass
        return {
            "file": path,
            "engine": "com",
            "excel_version": str(xl.app.Version),
            "project": str(vbp.Name),
            "protection": protection,
            "protected": protection != 0,
            "components": components,
            "references": references,
            "sha256": sha256_file(path),
        }
    finally:
        xl.close()


def _find_component(vbp, module: str):
    try:
        return vbp.VBComponents(module)
    except Exception as exc:
        names = []
        try:
            names = [str(c.Name) for c in vbp.VBComponents]
        except Exception as enum_exc:
            # A protected/unviewable project must not masquerade as a missing module.
            raise _map_com_error(enum_exc) from enum_exc
        raise BridgeError(
            "MODULE_NOT_FOUND",
            f"module not found: {module}",
            "available modules: " + ", ".join(names) if names else None,
        ) from exc


def com_read(path: str, module: str) -> dict:
    xl = Excel()
    try:
        xl.open(path, readonly=True)
        vbp = xl.project()
        comp = _find_component(vbp, module)
        cm = comp.CodeModule
        lines = int(cm.CountOfLines)
        code = cm.Lines(1, lines) if lines else ""
        return {
            "file": path,
            "engine": "com",
            "module": str(comp.Name),
            "type": _component_type(comp),
            "lines": lines,
            "code": code,
            "eol": "CRLF",
            "sha256_code": hashlib.sha256(norm_code(code).encode("utf-8")).hexdigest(),
        }
    finally:
        xl.close()


def _backup_path(path: str) -> str:
    stamp = time.strftime("%Y%m%d-%H%M%S")
    base, ext = os.path.splitext(path)
    candidate = f"{base}.{stamp}.bak{ext}"
    n = 1
    while os.path.exists(candidate):
        candidate = f"{base}.{stamp}-{n}.bak{ext}"
        n += 1
    return candidate


def com_write(
    path: str,
    module: str,
    code: str,
    mode: str = "replace",
    create: bool = False,
    backup: bool = True,
    verify: bool = True,
) -> dict:
    if mode not in ("replace", "append"):
        raise BridgeError("INVALID_REQUEST", "mode must be 'replace' or 'append'")
    lock = os.path.join(os.path.dirname(path), "~$" + os.path.basename(path))
    if os.path.exists(lock):
        raise BridgeError(
            "FILE_LOCKED",
            f"Excel lock file present: {lock}",
            "close the workbook in Excel (or delete a stale lock file) and retry",
        )

    backup_path = None
    if backup:
        backup_path = _backup_path(path)
        shutil.copy2(path, backup_path)

    before = sha256_file(path)
    verified = None
    xl = Excel()
    try:
        xl.open(path, readonly=False)
        vbp = xl.project()
        try:
            protection = int(getattr(vbp, "Protection", 0) or 0)
        except Exception:
            protection = 0
        if protection != 0:
            raise BridgeError(
                "PROJECT_PROTECTED",
                "the VBA project is password-protected; editing requires removing the password in the VBE",
            )
        created = False
        try:
            comp = vbp.VBComponents(module)
        except Exception as lookup_exc:
            names = []
            try:
                names = [str(c.Name) for c in vbp.VBComponents]
            except Exception as enum_exc:
                raise _map_com_error(enum_exc) from enum_exc
            if not create:
                raise BridgeError(
                    "MODULE_NOT_FOUND",
                    f"module not found: {module}",
                    "available modules: " + ", ".join(names),
                ) from lookup_exc
            try:
                comp = vbp.VBComponents.Add(1)  # vbext_ct_StdModule
                comp.Name = module
            except Exception as add_exc:
                raise _map_com_error(add_exc) from add_exc
            created = True

        cm = comp.CodeModule
        existing = cm.Lines(1, int(cm.CountOfLines)) if int(cm.CountOfLines) else ""
        if mode == "append" and existing.strip():
            merged = ensure_trailing_newline(existing) + "\r\n" + ensure_trailing_newline(code)
        else:
            merged = ensure_trailing_newline(code)

        # Rewrite the whole code area (hidden VB_* attributes are preserved by Excel).
        if int(cm.CountOfLines):
            cm.DeleteLines(1, int(cm.CountOfLines))
        cm.AddFromString(merged)
        lines_after = int(cm.CountOfLines)

        xl.wb.Save()
        after = sha256_file(path)
        result = {
            "file": path,
            "engine": "com",
            "module": str(comp.Name),
            "mode": mode,
            "created": created,
            "backup": backup_path,
            "changed": before != after,
            "lines": lines_after,
            "sha256_before": before,
            "sha256_after": after,
        }
        if verify:
            # Reuse the same Excel instance for the reopen check; a second
            # DispatchEx would cost several extra seconds per write.
            try:
                xl.wb.Close(SaveChanges=False)
                xl.wb = None
                xl.open(path, readonly=True)
                comp2 = _find_component(xl.project(), module)
                cm2 = comp2.CodeModule
                lines2 = int(cm2.CountOfLines)
                code2 = cm2.Lines(1, lines2) if lines2 else ""
                verified = {
                    "lines": lines2,
                    "sha256_code": hashlib.sha256(norm_code(code2).encode("utf-8")).hexdigest(),
                }
            except Exception as exc:
                result["verify_error"] = _com_error_text(exc)
    finally:
        xl.close()

    if verify:
        if verified is not None:
            result["verified_lines"] = verified["lines"]
            result["verified_sha256_code"] = verified["sha256_code"]
            result["code_sha256"] = hashlib.sha256(norm_code(merged).encode("utf-8")).hexdigest()
            result["verified"] = result["verified_sha256_code"] == result["code_sha256"]
        else:
            result["verified"] = False
    return result


# --------------------------------------------------------------------------- export


def _safe_filename(name: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name).strip().rstrip(". ")
    return cleaned or "Module"


def _write_text(path: str, text: str) -> None:
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(text)


def com_export(
    path: str,
    out_dir: str,
    mode: str = "code",
    include: list[str] | None = None,
    overwrite: bool = False,
) -> dict:
    """Export every component as source (mode=code) or as VBA export files (mode=component)."""
    if mode not in ("code", "component"):
        raise BridgeError("INVALID_REQUEST", "mode must be 'code' or 'component'")
    raw = str(out_dir or "").strip()
    if not raw:
        raise BridgeError("INVALID_REQUEST", "outDir is required")
    if not os.path.isabs(raw):
        raw = os.path.join(os.path.dirname(path), raw)
    out = os.path.normpath(os.path.abspath(raw))
    if os.path.exists(out) and not os.path.isdir(out):
        raise BridgeError("INVALID_OUTDIR", f"outDir is a file: {out}")
    os.makedirs(out, exist_ok=True)

    wanted = {str(name).lower() for name in include} if include else None
    xl = Excel()
    exported: list[dict] = []
    skipped: list[dict] = []
    errors: list[dict] = []
    try:
        xl.open(path, readonly=True)
        vbp = xl.project()
        try:
            comps = list(vbp.VBComponents)
        except Exception as exc:
            raise _map_com_error(exc) from exc
        for comp in comps:
            name = str(comp.Name)
            if wanted is not None and name.lower() not in wanted:
                continue
            type_id = int(comp.Type)
            ext = TYPE_EXT.get(type_id, ".txt")
            stem = _safe_filename(name)
            target = os.path.join(out, stem + ext)
            try:
                lines = int(comp.CodeModule.CountOfLines)
            except Exception:
                lines = 0
            if os.path.exists(target) and not overwrite:
                skipped.append({"module": name, "file": os.path.basename(target), "reason": "exists"})
                continue
            try:
                if os.path.exists(target):
                    os.remove(target)
                if mode == "code":
                    code = comp.CodeModule.Lines(1, lines) if lines else ""
                    _write_text(target, code)
                    files = [os.path.basename(target)]
                else:
                    comp.Export(target)
                    files = [os.path.basename(target)]
                    for entry in sorted(os.listdir(out)):
                        if entry != os.path.basename(target) and entry.startswith(stem + "."):
                            files.append(entry)
                if not os.path.exists(target):
                    raise RuntimeError("export produced no file")
                exported.append(
                    {
                        "module": name,
                        "type": TYPE_NAMES.get(type_id, f"type{type_id}"),
                        "type_id": type_id,
                        "file": os.path.basename(target),
                        "files": files,
                        "lines": lines,
                        "bytes": os.path.getsize(target),
                        "sha256": sha256_file(target),
                        "design_included": bool(mode == "component" or type_id != 3),
                    }
                )
            except Exception as exc:
                errors.append({"module": name, "error": _com_error_text(exc)})
    finally:
        xl.close()

    manifest = {
        "file": path,
        "engine": "com",
        "mode": mode,
        "outDir": out,
        "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "components": exported,
        "skipped": skipped,
        "errors": errors,
    }
    if mode == "code":
        manifest["notes"] = [
            "mode=code exports the code module only; UserForm .frm files contain code-behind "
            "without the form designer, use mode=component for .frm + .frx"
        ]
    else:
        manifest["notes"] = ["mode=component uses VBComponent.Export; forms include .frm + .frx"]
    manifest_path = os.path.join(out, "_manifest.json")
    _write_text(manifest_path, json.dumps(manifest, ensure_ascii=False, indent=2))
    return {
        "file": path,
        "engine": "com",
        "mode": mode,
        "outDir": out,
        "manifest": manifest_path,
        "exported": exported,
        "skipped": skipped,
        "errors": errors,
        "counts": {
            "exported": len(exported),
            "skipped": len(skipped),
            "errors": len(errors),
        },
    }


# ------------------------------------------------------------------------ procedures

PROC_START = re.compile(
    r"^\s*(?:(Public|Private|Friend)\s+)?(?:(Static)\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)",
    re.IGNORECASE,
)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)


def _is_comment(line: str) -> bool:
    stripped = line.lstrip()
    return stripped.startswith("'") or stripped.lower().startswith("rem ")


def _logical_line(lines: list[str], index: int) -> tuple[str, int]:
    """Join a declaration continued with trailing underscores; return text + last index."""
    text = lines[index]
    last = index
    while text.rstrip().endswith("_") and last + 1 < len(lines):
        last += 1
        text = text.rstrip()[:-1] + " " + lines[last].strip()
    return text, last


def parse_procedures(code: str) -> list[dict]:
    """Line numbers (1-based) for Sub/Function/Property declarations in a module."""
    lines = norm_code(code).split("\n")
    procs: list[dict] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line.strip() or _is_comment(line) or line.lstrip().lower().startswith("attribute "):
            i += 1
            continue
        match = PROC_START.match(line)
        if not match:
            i += 1
            continue
        scope = match.group(1) or "default"
        static = bool(match.group(2))
        kind = " ".join(word.capitalize() for word in match.group(3).split())
        name = match.group(4)
        signature, signature_end = _logical_line(lines, i)
        open_paren = signature.find("(")
        close_paren = signature.rfind(")")
        args = ""
        if 0 <= open_paren < close_paren:
            args = " ".join(signature[open_paren + 1 : close_paren].split())

        end = None
        j = signature_end + 1
        while j < len(lines):
            if PROC_END.match(lines[j]) and not _is_comment(lines[j]):
                end = j
                break
            if PROC_START.match(lines[j]) and not _is_comment(lines[j]):
                break
            j += 1
        procs.append(
            {
                "name": name,
                "kind": kind,
                "scope": scope,
                "static": static,
                "args": args,
                "start": i + 1,
                "end": (end + 1) if end is not None else len(lines),
            }
        )
        i = (end + 1) if end is not None else max(j, i + 1)
    return procs


def com_procedures(path: str, module: str | None = None) -> dict:
    xl = Excel()
    try:
        xl.open(path, readonly=True)
        vbp = xl.project()
        if module:
            comps = [_find_component(vbp, module)]
        else:
            try:
                comps = list(vbp.VBComponents)
            except Exception as exc:
                raise _map_com_error(exc) from exc
        modules = []
        total = 0
        for comp in comps:
            lines = int(comp.CodeModule.CountOfLines)
            code = comp.CodeModule.Lines(1, lines) if lines else ""
            procs = parse_procedures(code)
            total += len(procs)
            modules.append(
                {
                    "module": str(comp.Name),
                    "type": _component_type(comp),
                    "type_id": int(comp.Type),
                    "lines": lines,
                    "procedures": procs,
                }
            )
        return {"file": path, "engine": "com", "modules": modules, "total": total}
    finally:
        xl.close()


def ovba_procedures(path: str, module: str | None = None) -> dict:
    mods = ovba_modules(path)
    if module:
        mods = [m for m in mods if m["name"].lower() == module.lower()]
        if not mods:
            raise BridgeError("MODULE_NOT_FOUND", f"module not found: {module}")
    modules = []
    total = 0
    for m in mods:
        procs = parse_procedures(m["code"])
        total += len(procs)
        modules.append(
            {
                "module": m["name"],
                "type": m["type"],
                "type_id": None,
                "lines": m["lines"],
                "procedures": procs,
            }
        )
    return {
        "file": path,
        "engine": "ovba",
        "modules": modules,
        "total": total,
        "line_numbers_note": "offline source includes the hidden Attribute lines, so line numbers count them",
    }


# ------------------------------------------------------------------------- offline


def _load_oletools():
    if VENDOR not in sys.path:
        sys.path.insert(0, VENDOR)
    try:
        from oletools.olevba import VBA_Parser  # type: ignore
    except Exception as exc:  # pragma: no cover - environment dependent
        raise BridgeError(
            "OLETOOLS_MISSING",
            "oletools is not available: " + str(exc),
            "pip install oletools (or run the bundled vendor directory)",
        ) from exc
    return VBA_Parser


def ovba_modules(path: str) -> list[dict]:
    VBA_Parser = _load_oletools()
    parser = VBA_Parser(path)
    try:
        if not parser.detect_vba_macros():
            return []
        mods = []
        for _filename, stream, vba_filename, code in parser.extract_macros():
            name = os.path.splitext(vba_filename)[0]
            kind = "document" if vba_filename.lower().endswith(".cls") else "std"
            mods.append(
                {
                    "name": name,
                    "stream": "/".join(stream),
                    "type": kind,
                    "lines": norm_code(code).count("\n") + (0 if not code else 1),
                    "code": code,
                }
            )
        return mods
    finally:
        parser.close()


def ovba_inspect(path: str) -> dict:
    mods = ovba_modules(path)
    container = "zip" if zipfile.is_zipfile(path) else "ole"
    return {
        "file": path,
        "engine": "ovba",
        "container": container,
        "project": "VBAProject",
        "protection": None,
        "protected": None,
        "components": [
            {"name": m["name"], "type": m["type"], "type_id": None, "lines": m["lines"]}
            for m in mods
        ],
        "references": [],
        "sha256": sha256_file(path),
        "note": "offline read: attributes are included in code, p-code is not interpreted, component types are approximate",
    }


def ovba_read(path: str, module: str) -> dict:
    mods = ovba_modules(path)
    for m in mods:
        if m["name"] == module:
            return {
                "file": path,
                "engine": "ovba",
                "module": m["name"],
                "type": m["type"],
                "lines": m["lines"],
                "code": m["code"],
                "eol": "raw",
                "sha256_code": hashlib.sha256(norm_code(m["code"]).encode("utf-8")).hexdigest(),
            }
    raise BridgeError(
        "MODULE_NOT_FOUND",
        f"module not found: {module}",
        "available modules: " + ", ".join(m["name"] for m in mods) if mods else None,
    )


# --------------------------------------------------------------------------- doctor


def _registry_accessvbom() -> tuple[int | None, str]:
    try:
        import winreg
    except Exception as exc:  # pragma: no cover - non-Windows
        return None, f"winreg unavailable: {exc}"
    for version in ("16.0", "15.0", "14.0", "12.0"):
        sub = rf"Software\Microsoft\Office\{version}\Excel\Security"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub) as key:
                value, _ = winreg.QueryValueEx(key, "AccessVBOM")
                return int(value), f"HKCU\\{sub}\\AccessVBOM"
        except FileNotFoundError:
            continue
        except OSError as exc:
            return None, f"{sub}: {exc}"
    return None, "AccessVBOM not found under HKCU Office 12.0-16.0"


def com_doctor(file_path: str | None = None) -> dict:
    checks: list[dict] = []

    def add(name: str, ok: bool, detail: str, hint: str | None = None, critical: bool = False):
        checks.append({"name": name, "ok": ok, "detail": detail, "hint": hint, "critical": critical})

    add(
        "python",
        True,
        f"{platform.python_version()} ({platform.machine()}) at {sys.executable}",
        critical=True,
    )

    try:
        import win32com  # noqa: F401

        add("pywin32", True, "win32com import ok", critical=True)
    except Exception as exc:
        add("pywin32", False, f"win32com import failed: {exc}", "pip install pywin32", critical=True)

    excel = None
    try:
        excel = Excel()
        add("excel-com", True, f"Excel {excel.app.Version} started via COM", critical=True)
    except BridgeError as exc:
        add("excel-com", False, exc.message, exc.hint, critical=True)
    except Exception as exc:
        add("excel-com", False, f"unexpected COM error: {exc}", critical=True)
    finally:
        if excel is not None:
            excel.close()

    access, access_detail = _registry_accessvbom()
    if access == 1:
        add("AccessVBOM", True, access_detail, critical=True)
    elif access == 0:
        add(
            "AccessVBOM",
            False,
            "AccessVBOM = 0 (Trust Center blocks programmatic VBA access)",
            "Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings > "
            'enable "Trust access to the VBA project object model"',
            critical=True,
        )
    else:
        add(
            "AccessVBOM",
            False,
            access_detail,
            "open Excel once and enable the Trust Center setting for VBA project object model access",
            critical=True,
        )

    try:
        _load_oletools()
        add("oletools", True, "offline OVBA parser import ok", critical=False)
    except Exception as exc:
        add("oletools", False, f"offline parser unavailable: {exc}", "reinstall the bundle", critical=False)

    try:
        fd, tmp = tempfile.mkstemp(prefix="vba-bridge-", suffix=".tmp")
        os.close(fd)
        with open(tmp, "w", encoding="utf-8") as handle:
            handle.write("ok")
        os.remove(tmp)
        add("temp-dir", True, "temporary directory is writable", critical=False)
    except Exception as exc:
        add("temp-dir", False, f"temporary directory is not writable: {exc}", critical=False)

    if file_path:
        lock = os.path.join(os.path.dirname(file_path), "~$" + os.path.basename(file_path))
        add("file-lock", not os.path.exists(lock), f"lock file: {lock}" if os.path.exists(lock) else "no Excel lock file")
        xl = None
        try:
            xl = Excel()
            xl.open(file_path, readonly=True)
            vbp = xl.project()
            components = list(vbp.VBComponents)
            protection = 0
            try:
                protection = int(vbp.Protection)
            except Exception:
                pass
            add(
                "workbook",
                True,
                f"{len(components)} component(s), protection={protection}",
                critical=False,
            )
        except BridgeError as exc:
            add("workbook", False, exc.message, exc.hint, critical=False)
        except Exception as exc:
            add("workbook", False, f"cannot inspect workbook: {exc}", critical=False)
        finally:
            if xl is not None:
                xl.close()

    ready = all(check["ok"] for check in checks if check["critical"])
    return {
        "file": file_path,
        "ready": ready,
        "checks": checks,
        "vendor": VENDOR,
        "python": sys.executable,
    }


# --------------------------------------------------------------------------- dispatch


def _engine_error_is_fallback(exc: BridgeError) -> bool:
    return exc.code in {"PYWIN32_MISSING", "EXCEL_UNAVAILABLE"}


def _wants_ovba(engine: str, path: str) -> bool:
    if engine == "ovba":
        return True
    return engine == "auto" and os.environ.get("VBA_BRIDGE_ENGINE", "").lower() == "ovba"


def dispatch(req: dict) -> dict:
    action = str(req.get("action") or "").lower()
    engine = str(req.get("engine") or "auto").lower()
    raw_file = str(req.get("file") or "")
    if action != "doctor" and not raw_file:
        raise BridgeError("INVALID_REQUEST", "file is required")
    path = validate_file(raw_file) if raw_file else None
    if engine not in {"auto", "com", "ovba"}:
        raise BridgeError("INVALID_REQUEST", "engine must be auto, com or ovba")

    if action == "doctor":
        return com_doctor(path)

    if action == "inspect":
        if _wants_ovba(engine, path):
            return ovba_inspect(path)
        try:
            return com_inspect(path)
        except BridgeError as exc:
            if engine == "auto" and exc.code in {"PYWIN32_MISSING", "EXCEL_UNAVAILABLE", "VBA_TRUST_DISABLED"}:
                data = ovba_inspect(path)
                data["warnings"] = [f"COM unavailable ({exc.code}); used offline parser"]
                data["com_error"] = {"code": exc.code, "message": exc.message, "hint": exc.hint}
                return data
            raise

    if action == "read":
        module = str(req.get("module") or "")
        if not module:
            raise BridgeError("INVALID_REQUEST", "module is required")
        if _wants_ovba(engine, path):
            return ovba_read(path, module)
        try:
            return com_read(path, module)
        except BridgeError as exc:
            if engine == "auto" and exc.code in {"PYWIN32_MISSING", "EXCEL_UNAVAILABLE", "VBA_TRUST_DISABLED"}:
                data = ovba_read(path, module)
                data["warnings"] = [f"COM unavailable ({exc.code}); used offline parser"]
                return data
            raise

    if action == "write":
        if engine == "ovba":
            raise BridgeError(
                "ENGINE_READONLY",
                "the offline OVBA engine is read-only in this build",
                "use the COM engine (Excel installed) to write",
            )
        module = str(req.get("module") or "")
        if not module:
            raise BridgeError("INVALID_REQUEST", "module is required")
        code = req.get("code")
        if code is None:
            raise BridgeError("INVALID_REQUEST", "code is required")
        return com_write(
            path,
            module,
            str(code),
            mode=str(req.get("mode") or "replace"),
            create=bool(req.get("create", False)),
            backup=bool(req.get("backup", True)),
            verify=bool(req.get("verify", True)),
        )

    if action == "export":
        if engine == "ovba":
            raise BridgeError(
                "ENGINE_READONLY",
                "the offline OVBA engine cannot export a project in this build",
                "use the COM engine (Excel installed) to export modules",
            )
        out_dir = str(req.get("outDir") or "")
        if not out_dir:
            raise BridgeError("INVALID_REQUEST", "outDir is required")
        include = req.get("include")
        if include is not None and not isinstance(include, list):
            raise BridgeError("INVALID_REQUEST", "include must be an array of module names")
        return com_export(
            path,
            out_dir,
            mode=str(req.get("mode") or "code"),
            include=[str(item) for item in include] if include else None,
            overwrite=bool(req.get("overwrite", False)),
        )

    if action == "procedures":
        module = str(req.get("module") or "") or None
        if _wants_ovba(engine, path):
            return ovba_procedures(path, module)
        try:
            return com_procedures(path, module)
        except BridgeError as exc:
            if engine == "auto" and exc.code in {"PYWIN32_MISSING", "EXCEL_UNAVAILABLE", "VBA_TRUST_DISABLED"}:
                data = ovba_procedures(path, module)
                data["warnings"] = [f"COM unavailable ({exc.code}); used offline parser"]
                return data
            raise

    raise BridgeError("INVALID_REQUEST", f"unknown action: {action!r}")


def _response(fn) -> dict:
    try:
        data = fn()
        return {"ok": True, "data": data}
    except BridgeError as exc:
        return {
            "ok": False,
            "error": {"code": exc.code, "message": exc.message, "hint": exc.hint},
        }
    except Exception as exc:  # pragma: no cover - defensive
        import traceback

        return {
            "ok": False,
            "error": {
                "code": "INTERNAL",
                "message": repr(exc),
                "hint": traceback.format_exc()[-2000:],
            },
        }


def main(argv: list[str]) -> int:
    for stream in (sys.stdin, sys.stdout):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass

    if len(argv) > 1 and argv[1] == "--json-stdin":
        try:
            req = json.loads(sys.stdin.read())
        except Exception as exc:
            print(json.dumps({"ok": False, "error": {"code": "BAD_JSON", "message": str(exc)}}))
            return 2
        print(json.dumps(_response(lambda: dispatch(req)), ensure_ascii=False))
        return 0

    parser = argparse.ArgumentParser(prog="vba_bridge")
    sub = parser.add_subparsers(dest="action", required=True)
    p_inspect = sub.add_parser("inspect")
    p_inspect.add_argument("file")
    p_inspect.add_argument("--engine", default="auto", choices=["auto", "com", "ovba"])
    p_read = sub.add_parser("read")
    p_read.add_argument("file")
    p_read.add_argument("module")
    p_read.add_argument("--engine", default="auto", choices=["auto", "com", "ovba"])
    p_write = sub.add_parser("write")
    p_write.add_argument("file")
    p_write.add_argument("module")
    p_write.add_argument("--code-file", required=True)
    p_write.add_argument("--mode", default="replace", choices=["replace", "append"])
    p_write.add_argument("--create", action="store_true")
    p_write.add_argument("--no-backup", action="store_true")
    p_write.add_argument("--no-verify", action="store_true")
    p_export = sub.add_parser("export")
    p_export.add_argument("file")
    p_export.add_argument("out_dir")
    p_export.add_argument("--mode", default="code", choices=["code", "component"])
    p_export.add_argument("--include", nargs="*")
    p_export.add_argument("--overwrite", action="store_true")
    p_proc = sub.add_parser("procedures")
    p_proc.add_argument("file")
    p_proc.add_argument("module", nargs="?")
    p_proc.add_argument("--engine", default="auto", choices=["auto", "com", "ovba"])
    args = parser.parse_args(argv[1:])

    if args.action == "inspect":
        req = {"action": "inspect", "file": args.file, "engine": args.engine}
    elif args.action == "read":
        req = {"action": "read", "file": args.file, "module": args.module, "engine": args.engine}
    elif args.action == "export":
        req = {
            "action": "export",
            "file": args.file,
            "outDir": args.out_dir,
            "mode": args.mode,
            "include": args.include,
            "overwrite": args.overwrite,
        }
    elif args.action == "procedures":
        req = {"action": "procedures", "file": args.file, "module": args.module, "engine": args.engine}
    else:
        with open(args.code_file, "r", encoding="utf-8-sig") as f:
            code = f.read()
        req = {
            "action": "write",
            "file": args.file,
            "module": args.module,
            "code": code,
            "mode": args.mode,
            "create": args.create,
            "backup": not args.no_backup,
            "verify": not args.no_verify,
            "engine": "com",
        }
    print(json.dumps(_response(lambda: dispatch(req)), ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
