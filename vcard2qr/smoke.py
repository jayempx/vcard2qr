"""Verify the packaged UI and its dependencies without extra test-only libraries."""

import json
import time
from pathlib import Path

from openpyxl import Workbook
from PIL import Image
from PySide6.QtCore import qVersion

from .qr import save_png


def run(app, window, output: Path) -> int:
    """Use synthetic contacts in a new directory and emit a machine-readable result."""
    # Refuse an existing directory so a smoke run cannot replace user data.
    output.mkdir(parents=True, exist_ok=False)

    def wait():
        deadline = time.monotonic() + 30
        while window.worker is not None and time.monotonic() < deadline:
            app.processEvents()
            time.sleep(0.005)
        if window.worker is not None:
            window.cancel_job()
            window.worker.wait()
            app.processEvents()
            raise RuntimeError("Background job timed out.")

    try:
        window.fields["First Name"].setText("Package")
        window.fields["Last Name"].setText("Check")
        window.fields["Email"].setText("check@example.com")
        window.generate_qr()
        if window.current_qr is None:
            raise RuntimeError(window.status.text())
        image = window.current_qr
        save_png(image, output / "single.png")
        window.copy_to_clipboard()
        clipboard_available = app.clipboard().image().width() == 512
        if not clipboard_available and "Clipboard unavailable" not in window.status.text():
            raise RuntimeError("Clipboard failure was not reported to the user.")
        with Image.open(output / "single.png") as exported:
            if exported.tobytes() != image.tobytes():
                raise RuntimeError("PNG data changed during export.")
        workbook = Workbook()
        workbook.active.append(["Nome", "Email"])
        workbook.active.append(["Example One", "one@example.com"])
        workbook.active.append(["Example Two", "two@example.com"])
        path = output / "contacts.xlsx"
        workbook.save(path)
        workbook.close()
        window.load_workbook_path(str(path))
        wait()
        if window.book is None:
            raise RuntimeError(window.batch_summary.text())
        window.export_to(str(output), ["First Name", "Email"])
        wait()
        if len(list(output.glob("vcard2qr-*/*.png"))) != 2:
            raise RuntimeError(window.batch_summary.text())
        if window.current_qr is not image:
            raise RuntimeError("Batch conversion replaced the preview.")
        window.fields["First Name"].setText("Changed")
        if window.save_button.isEnabled():
            raise RuntimeError("Stale image export remains enabled.")
        result = {
            "passed": True,
            "qt": qVersion(),
            "batch_saved": 2,
            "preview_preserved": True,
            "clipboard_available": clipboard_available,
        }
        code = 0
    except Exception as exc:
        result, code = {"passed": False, "error": f"{type(exc).__name__}: {exc}"}, 1
    finally:
        window.close()
    (output / "result.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    return code
