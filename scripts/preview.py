"""Capture repeatable documentation screenshots using synthetic contact details."""

import argparse
import os
import tempfile
import time
from pathlib import Path

# Offscreen rendering keeps documentation generation independent of the desktop.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook
from PySide6.QtGui import QFontDatabase
from PySide6.QtWidgets import QApplication

from vcard2qr.app import VCardQRApp


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    app = QApplication([])
    # Windows offscreen Qt does not discover system fonts automatically.
    if os.name == "nt":
        for name in ("segoeui.ttf", "segoeuib.ttf"):
            QFontDatabase.addApplicationFont(str(Path(os.environ["WINDIR"]) / "Fonts" / name))
    app.setStyle("Fusion")
    window = VCardQRApp()
    window.fields["First Name"].setText("Alex")
    window.fields["Last Name"].setText("Morgan")
    window.fields["Organization"].setText("Studio Example")
    window.fields["Title"].setText("Creative producer")
    window.fields["Email"].setText("alex@example.com")
    window.fields["Mobile"].setText("+39 333 0123456")
    window.fields["Address"].setText("Via Esempio 12, Bologna")
    window.generate_qr()
    window.show()
    for theme in ("Light", "Dark"):
        window.theme.setCurrentText(theme)
        app.processEvents()
        if not window.grab().save(str(args.output / f"contact-{theme.lower()}.webp"), "WEBP"):
            raise RuntimeError("Could not write the preview image.")
    # Exercise the real asynchronous workbook preview with synthetic rows.
    with tempfile.TemporaryDirectory(prefix="vcard2qr-preview-") as temporary:
        path = Path(temporary) / "example-contacts.xlsx"
        book = Workbook()
        book.active.append(["Nome", "Cognome", "Email", "Azienda"])
        book.active.append(["Alex", "Morgan", "alex@example.com", "Studio Example"])
        book.active.append(["Jamie", "River", "jamie@example.com", "Studio Example"])
        book.save(path)
        book.close()
        window.tabs.setCurrentIndex(1)
        window.theme.setCurrentText("Light")
        window.load_workbook_path(str(path))
        deadline = time.monotonic() + 15
        while window.worker is not None and time.monotonic() < deadline:
            app.processEvents()
            time.sleep(0.01)
        if window.worker is not None:
            window.cancel_job()
            window.worker.wait()
            app.processEvents()
            raise RuntimeError("Workbook preview did not finish in time.")
        app.processEvents()
        window.grab().save(str(args.output / "excel-import.webp"), "WEBP")
    window.resize(880, 650)
    app.processEvents()
    window.grab().save(str(args.output / "compact-layout.webp"), "WEBP")
    window.close()


if __name__ == "__main__":
    main()
