# Third-party notices

This package bundles a minimal subset of the following Python packages for the
offline VBA parser. They are **not** installed into the user's Python
environment; they are imported from `vendor/`.

| Package | Version | License | Copyright |
| --- | --- | --- | --- |
| [oletools](https://github.com/decalage2/oletools) | 0.60.2 | BSD-2-Clause | Copyright (c) 2012-2021 Philippe Lagadec / DECALAGE |
| [olefile](https://github.com/decalage2/olefile) | 0.47 | BSD-2-Clause | Copyright (c) 2005-2023 Philippe Lagadec |
| [colorclass](https://github.com/matthewdeanmartin/colorclass) | 2.2.2 | MIT | Copyright (c) 2012-2016 Mattias Lind |
| [pyparsing](https://github.com/pyparsing/pyparsing) | 3.3.3 | MIT | Copyright (c) 2003-2024 Paul T. McGuire |

The oletools subset also carries its bundled `prettytable` (BSD) and `xglob`
(BSD) helpers, and `oledump`'s `plugin_biff.py` (public domain by Didier
Stevens).

Full license texts for oletools and olefile are included in
`vendor/oletools/LICENSE.txt` and `vendor/olefile/LICENSE.txt`; colorclass and
pyparsing license texts are available in their upstream repositories linked
above.

The bridge itself does not bundle `cryptography`, `msoffcrypto-tool`,
`pcodedmp`, `win-unicode-console`, or `xxxswf`; encrypted-file decryption,
p-code disassembly and Flash extraction are outside this plugin's scope. The
GPL-licensed `pcodedmp` and `xxxswf` modules were removed from the vendored
subset on purpose.

The Windows COM bridge uses the user's own Python installation with `pywin32`
(PSF-2.0). Microsoft Excel is a commercial product and is not distributed with
this package.
