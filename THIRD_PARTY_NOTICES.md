# Third-party notices

The project owner has not selected a license for vcard2qr. This file records dependencies; it does not license the application or replace the dependencies' complete license texts.

| Runtime component | Upstream and license information |
|---|---|
| Python | [Python license](https://docs.python.org/3/license.html) |
| PySide6-Essentials / Shiboken / Qt | [Qt for Python licenses](https://doc.qt.io/qtforpython-6/licenses.html); inspect each included component's notices. |
| qrcode | [Upstream repository](https://github.com/lincolnloop/python-qrcode), BSD license. |
| Pillow | [Pillow license and bundled libraries](https://pillow.readthedocs.io/en/stable/about.html#license). |
| openpyxl | [Package metadata](https://pypi.org/project/openpyxl/), MIT license. |
| et-xmlfile | [Package metadata](https://pypi.org/project/et-xmlfile/), MIT license. |
| colorama | [Upstream repository](https://github.com/tartley/colorama), BSD license. |

The Windows packaging script copies installed runtime distribution metadata and license files into the application bundle and includes this notice. PyInstaller remains a build-time tool. Keep the complete bundle together when testing or distributing it.

Before a public release, choose a project license and review the complete notices and distribution obligations for the actual bundled binaries. Dependency license files remain authoritative.
