# Changelog

## 0.1.1 (2026-10-19)

- Use an absolute GitHub raw URL for the README screenshot so it renders on
  the npm package page.

## 0.1.0 (2026-10-18)

First public release.

- `vba_inspect` — project name, protection, components and references.
- `vba_read_module` — read code from standard, class, document and UserForm modules.
- `vba_write_module` — replace/append/create modules, save in place, reopen and
  verify by SHA-256, timestamped `.bak` backup.
- `vba_export_modules` — export all components as source (`mode=code`) or as
  native VBA files with `.frm` + `.frx` for forms (`mode=component`), plus
  `_manifest.json`.
- `vba_list_procedures` — Sub/Function/Property declarations with arguments and
  start/end lines, including UserForm event handlers.
- `vba_doctor` — preflight checks for Python, pywin32, Excel COM, `AccessVBOM`,
  the offline parser, temp directory and an optional workbook.
- Relative paths resolve against the DSH session working directory.
- Writes verify by reopening in the same Excel instance (about 5 s instead of
  paying a second Excel startup).
- Offline read fallback with a minimal bundled oletools subset (no GPL modules).
- Windows + desktop Excel + `AccessVBOM=1` required for COM operations.
