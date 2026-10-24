# Architecture and maintenance

The application remains a local desktop tool. There is no web server, database, telemetry or runtime network dependency.

| Module | Responsibility |
|---|---|
| `vcard.py` | Stable field keys, validation, escaping, vCard 3.0 serialization and byte-aware folding. |
| `qr.py` | Immutable appearance settings, contrast checks, QR rasterization and atomic PNG saving. |
| `excel.py` | Header suggestions, mapping validation, bounded preview, streamed batch export and row reports. |
| `app.py` | Qt widgets, themes, contact snapshots, worker lifecycle and native dialogs. |
| `__main__.py` | Installed module entry point. |
| `vcard-2-qr.py` | Compatibility launcher for existing shortcuts and packaging. |

## Data flow

The form supplies a mapping and optional custom entries to `build_vcard`. It returns a CRLF-delimited UTF-8-compatible string. `render_qr` returns a Pillow image. The UI copies this into an owned `QImage` and only then enables exports. Any contact or appearance edit clears the snapshot and disables both export actions.

The Excel path calls the same serializer and renderer, avoiding separate single/batch implementations. The worker receives immutable copies of settings, headers and mappings; it does not read or change widgets. Signals return progress, completion or failure to the main thread. Completed batch images never become the single-contact image.

## Format decisions

- Keep vCard **3.0** for continuity with existing output. Always include `VERSION`, `FN` and structured `N`. If personal names are absent, use organization, email or phone as `FN` and retain an empty structured `N`.
- Escape backslashes, newlines, commas and semicolons in text values. Validate phone and URI values separately. Reject syntax in extension names.
- Fold logical lines at 75 UTF-8 octets without splitting code points; continuation whitespace counts toward the limit. Use CRLF and a final terminator.
- Use `TYPE=WORK,VOICE`, not the old malformed `TYPE=WORK;VOICE`.
- Treat the single address field as the street-address component. Splitting city, postcode and country into dedicated inputs is outside this migration.

## Rendering decisions

`qrcode.get_matrix()` already includes its configured border. Render that matrix once, with integer module dimensions, centered in the requested canvas. Reject fewer than three pixels per module. Do not resize the exported raster with Lanczos or another interpolation filter.

The optional radius is capped at one quarter of the module size. Finder patterns remain square. A dark-on-light 4.5:1 contrast check is a conservative UI guard, not a QR specification guarantee. Transparent output is checked against white; actual scanning depends on the final background.

`qrcode` 8.2 can signal capacity overflow as either `DataOverflowError` or `ValueError` for version 41. Both are translated at the library boundary. Exact payload recovery is tested with ZXing, which is a development-only dependency.

## Dependencies and checks

Use `requirements.lock` for runtime reproduction and `requirements-dev.lock` for tests/builds. Direct dependencies also appear in `pyproject.toml`; update these together. Both lock files pin versions, but they are not cryptographic hash locks. Test dependency updates in a fresh environment before changing the pins.

PySide6-Essentials provides the required Core, Gui and Widgets modules without Qt WebEngine. No Qt 5 dependency remains. Run `python -m pytest -q`, `python -m ruff check .` and `python -m ruff format --check .` from the checkout. Tests do not require a running desktop.

Generate README screenshots with `python -m scripts.preview --output docs/images`. This uses synthetic data. On Windows, the preview script explicitly registers Segoe UI because the offscreen Qt platform may otherwise draw missing-glyph boxes. The real app uses normal platform font discovery.

Build a Windows directory with `scripts/build_windows.ps1`. Build output belongs outside the repository. Do not commit environments, caches, local contact exports or build output. The existing `.gitignore` has intentionally been left unchanged.

The build script temporarily restricts `PATH` to the chosen Python environment and Windows system directories, then restores it. This prevents PyInstaller from collecting incompatible ICU/UCRT DLLs from tools such as Poppler. A clean analysis is required after changing DLL search paths. Diagnose frozen import errors with the selected paths in `Analysis-00.toc`; do not assume a missing-module warning proves the cause.

The packaged app supports `--smoke-test OUTPUT`, where `OUTPUT` must be a new directory. This exercises generation, PNG saving, clipboard conversion, Excel preview/export and stale-state invalidation with synthetic data, writes `result.json`, and exits. Set `QT_QPA_PLATFORM=offscreen` for headless execution. Run this on the built executable as well as an installed wheel; a successful source test does not prove a portable bundle works.

The smoke result reports `clipboard_available` separately. A restricted Windows session may deny clipboard access; in that case the test verifies the UI reports the failure and continues testing export. The offscreen CI smoke check requires clipboard availability. A native session with access denied still needs an interactive clipboard test before claiming full clipboard validation.

## Making changes

Keep comments and docstrings in English. Add comments around format rules, ownership boundaries and non-obvious failure handling. When changing serialization or rendering, add a regression that decodes the exported PNG. When changing imports, exercise a malformed row between two valid rows and cancellation after a completed export. When changing UI state, verify that copied/saved output matches the visible preview and that closing a worker cannot destroy a running thread.
