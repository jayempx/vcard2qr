# Validation record

Local verification on 2026-10-24, starting from upstream commit `86e72ad620ed9c72f7e4717845db208c3e9a752d`. This record describes pre-publication local verification; no GitHub release or remote CI execution is implied.

Environment: Windows 11 x64, Python 3.14.5, PySide6-Essentials/Qt 6.12.0, qrcode 8.2, Pillow 13.0.0, openpyxl 3.1.5. Test/build packages are recorded in `requirements-dev.lock`.

## Completed checks

| Check | Result |
|---|---|
| Fresh dependency installation and `pip check` | Passed using only Qt Essentials and the locked dependency set. |
| Regression suite | 33 tests passed. |
| Code style | Ruff lint and format checks passed. |
| vCard serialization | Required names, typed values, delimiter escaping, CRLF and multibyte line folding verified. |
| QR decoding | ZXing decoded exported PNGs and recovered the original UTF-8 payload, including transparency on white and rounded modules. |
| Excel import | Supplied sample imports two contacts; malformed rows, formulas, collisions, cancellation and repeat exports covered. |
| UI state | Editing invalidates exports; batch work preserves the single preview; errors restore controls; worker shutdown is cooperative. |
| Visual review | Light, dark, Excel preview and compact layouts rendered and inspected. |
| Python distributions | Source archive and wheel built; source archive contains fixtures, sample, scripts and documentation. |
| Portable executable | Built with PyInstaller; offscreen smoke run passed with clipboard readback available. |
| Native Windows Qt backend | Generation, PNG save, workbook preview and batch export passed. Clipboard access is unavailable in this execution session. |

The smoke command uses synthetic contacts and writes `result.json` plus one single-contact PNG and two batch PNGs. The images produced by the executable are independently decoded outside the bundle. The app does not ship ZXing as a runtime dependency.

## Packaging issue resolved

The first frozen build failed importing `QtCore`, while the same virtual environment ran the source correctly. Inspection of `Analysis-00.toc` showed an incompatible `icuuc.dll` collected from an unrelated Poppler installation on `PATH`. A diagnostic replacement confirmed the cause. The final build script fixes dependency resolution by isolating `PATH` and rebuilding from a clean analysis; it does not rely on editing the finished bundle.

## Clipboard boundary

The native Windows session returned access denied (`OpenClipboard` error 5), both with the source application and the executable. Offscreen clipboard conversion and readback pass. A dedicated regression verifies that silent clipboard failure is reported to the user rather than shown as success. Native interactive clipboard behavior remains a manual check.

## Before a public release

- Run the committed GitHub workflows; they are configured but have not run remotely for this working tree.
- Scan exported QR codes on the intended Android/iOS devices and confirm contact field import.
- Exercise copy/paste in an interactive Windows desktop session with clipboard permission.
- Test the portable folder on another Windows machine without Python installed.
- Check high-DPI behavior on the intended displays and any additional supported operating systems.
- Choose the project's license, review bundled notices and decide whether to sign the Windows executable.

## Reproduce

```powershell
# Run source validation from the repository root.
python -m pip install -r requirements-dev.lock
python -m pytest -q
python -m ruff check .
python -m ruff format --check .

# Verify a built executable; choose an output directory that does not already exist.
$env:QT_QPA_PLATFORM = 'offscreen'
$process = Start-Process -FilePath './vcard2qr.exe' -ArgumentList @('--smoke-test', 'C:/temp/vcard2qr-check') -WindowStyle Hidden -PassThru
$process.WaitForExit()
Get-Content C:/temp/vcard2qr-check/result.json
```

For offscreen verification, require both `passed: true` and `clipboard_available: true`. For native runs, report clipboard availability separately; a passing export check is not proof that clipboard access was available.
