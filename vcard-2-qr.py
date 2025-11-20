"""
vcard2qr GUI
===========================================
A PyQt5 application that lets you fill out contact information, generate a
vCard 3.0 payload, and render it as a customizable QR code.

Dependencies (install with pip):
    pip install PyQt5 qrcode[pil] pillow

Author: ChatGPT (OpenAI o3)
Date: 2025‑08‑06
"""
#!/usr/bin/env python

import sys
from typing import List, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QFormLayout,
    QHBoxLayout,
    QVBoxLayout,
    QLineEdit,
    QPushButton,
    QLabel,
    QColorDialog,
    QSpinBox,
    QScrollArea,
    QFileDialog,
    QMessageBox,
    QCheckBox,
)

import qrcode
from qrcode.constants import ERROR_CORRECT_Q
from PIL import Image, ImageDraw


class VCardQRApp(QWidget):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("vcard2qr")

        # ---------- Form fields ----------
        self.fields = {
            "First Name": QLineEdit(),
            "Last Name": QLineEdit(),
            "Organization": QLineEdit(),
            "Title": QLineEdit(),
            "Email": QLineEdit(),
            "Mobile": QLineEdit(),
            "Switchboard": QLineEdit(),
            "Direct Office": QLineEdit(),
            "Address": QLineEdit(),
            "LinkedIn": QLineEdit(),
        }
        self.custom_fields: List[Tuple[QLineEdit, QLineEdit]] = []

        # ---------- Layout scaffolding ----------
        root = QHBoxLayout(self)

        # -- Left‑hand scrollable form
        form_container = QWidget()
        self.form_layout = QFormLayout(form_container)
        for label, widget in self.fields.items():
            self.form_layout.addRow(label, widget)

        add_custom_btn = QPushButton("Add Custom Field")
        add_custom_btn.clicked.connect(self.add_custom_field)
        self.form_layout.addRow(add_custom_btn)

        scroll = QScrollArea()
        scroll.setWidget(form_container)
        scroll.setWidgetResizable(True)
        root.addWidget(scroll, stretch=2)

        # -- Right‑hand controls & preview
        ctrl_col = QVBoxLayout()

        # Size selector
        ctrl_col.addWidget(QLabel("Size (px):"))
        self.size_spin = QSpinBox(minimum=128, maximum=2048, value=512) # type: ignore
        ctrl_col.addWidget(self.size_spin)

        # Corner radius
        ctrl_col.addWidget(QLabel("Corner radius (px):"))
        self.radius_spin = QSpinBox(minimum=0, maximum=20, value=4) # type: ignore
        ctrl_col.addWidget(self.radius_spin)

        # Colors
        self.fg_color = "#000000"
        self.bg_color = "#FFFFFF"
        self.transparent_bg = False

        fg_btn = QPushButton("Foreground color …")
        fg_btn.clicked.connect(self.pick_fg_color)
        ctrl_col.addWidget(fg_btn)

        self.bg_btn = QPushButton("Background color …")
        self.bg_btn.clicked.connect(self.pick_bg_color)
        ctrl_col.addWidget(self.bg_btn)

        self.transparent_checkbox = QCheckBox("Sfondo trasparente")
        self.transparent_checkbox.toggled.connect(self.toggle_transparent_background)
        ctrl_col.addWidget(self.transparent_checkbox)

        # Generate / actions
        gen_btn = QPushButton("Generate QR")
        gen_btn.clicked.connect(self.generate_qr)
        ctrl_col.addWidget(gen_btn)

        copy_btn = QPushButton("Copy to clipboard")
        copy_btn.clicked.connect(self.copy_to_clipboard)
        ctrl_col.addWidget(copy_btn)

        save_btn = QPushButton("Save as PNG …")
        save_btn.clicked.connect(self.save_png)
        ctrl_col.addWidget(save_btn)

        ctrl_col.addStretch(1)

        # QR preview
        self.qr_label = QLabel(alignment=Qt.AlignCenter) # type: ignore
        self.qr_label.setMinimumSize(256, 256)
        ctrl_col.addWidget(self.qr_label, stretch=2)

        root.addLayout(ctrl_col, stretch=1)

        self.current_qr: Image.Image | None = None

    # ---------------------------------------------------------------------
    # Utility helpers
    # ---------------------------------------------------------------------
    def add_custom_field(self) -> None:
        """Append a new key/value pair to the form."""
        key_edit, val_edit = QLineEdit(), QLineEdit()
        self.custom_fields.append((key_edit, val_edit))
        self.form_layout.addRow(key_edit, val_edit)

    def pick_fg_color(self) -> None:
        color = QColorDialog.getColor()
        if color.isValid():
            self.fg_color = color.name()

    def pick_bg_color(self) -> None:
        color = QColorDialog.getColor()
        if color.isValid():
            self.bg_color = color.name()

    def toggle_transparent_background(self, checked: bool) -> None:
        """Toggle whether the QR background should be transparent."""
        self.transparent_bg = checked
        self.bg_btn.setEnabled(not checked)

    # ---------------------------------------------------------------------
    # vCard & QR generation
    # ---------------------------------------------------------------------
    def build_vcard(self) -> str:
        """Compose a vCard 3.0 string from the form contents."""
        lines: List[str] = ["BEGIN:VCARD", "VERSION:3.0"]

        fn = f"{self.fields['First Name'].text()} {self.fields['Last Name'].text()}".strip()
        if fn:
            lines.append(f"FN:{fn}")

        if org := self.fields["Organization"].text():
            lines.append(f"ORG:{org}")
        if title := self.fields["Title"].text():
            lines.append(f"TITLE:{title}")

        tel_map = {
            "Mobile": "CELL",
            "Switchboard": "WORK",
            "Direct Office": "WORK;VOICE",
        }
        for key, vtype in tel_map.items():
            if num := self.fields[key].text():
                lines.append(f"TEL;TYPE={vtype}:{num}")

        if email := self.fields["Email"].text():
            lines.append(f"EMAIL;TYPE=INTERNET:{email}")

        if addr := self.fields["Address"].text():
            # Basic ADR format: ADR;TYPE=WORK:;;street;city;region;zip;country
            lines.append(f"ADR;TYPE=WORK:;;{addr};;;;")

        if url := self.fields["LinkedIn"].text():
            lines.append(f"URL:{url}")

        # Custom data as X‑extensions
        for key_edit, val_edit in self.custom_fields:
            key = key_edit.text().strip().upper().replace(" ", "_")
            val = val_edit.text().strip()
            if key and val:
                lines.append(f"X-{key}:{val}")

        lines.append("END:VCARD")
        return "\n".join(lines)

    def generate_qr(self) -> None:
        vcard = self.build_vcard()
        if not vcard or vcard == "BEGIN:VCARD\nVERSION:3.0\nEND:VCARD":
            QMessageBox.warning(self, "Incomplete", "Add at least one field first.")
            return

        qr = qrcode.QRCode(
            version=None,
            error_correction=ERROR_CORRECT_Q,
            box_size=10,
            border=4,
        )
        qr.add_data(vcard)
        qr.make(fit=True)

        img = self.render_rounded(
            qr,
            self.size_spin.value(),
            self.radius_spin.value(),
            self.transparent_bg,
        )
        self.current_qr = img

        # Display in Qt label
        self.show_qr(img)

    def render_rounded(
        self,
        qr: qrcode.QRCode,
        size: int,
        radius: int,
        transparent_bg: bool,
    ) -> Image.Image:
        """Return a Pillow image with optional rounded modules."""
        matrix = qr.get_matrix()
        modules = len(matrix)
        box = size // (modules + 8)  # +8 for border 4 each side
        img_size = box * (modules + 8)
        mode = "RGBA" if transparent_bg else "RGB"
        background = (0, 0, 0, 0) if transparent_bg else self.bg_color
        img = Image.new(mode, (img_size, img_size), background)
        draw = ImageDraw.Draw(img)

        for r, row in enumerate(matrix):
            for c, val in enumerate(row):
                if val:
                    x0 = (c + 4) * box
                    y0 = (r + 4) * box
                    x1 = x0 + box
                    y1 = y0 + box
                    if radius:
                        draw.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=self.fg_color)
                    else:
                        draw.rectangle([x0, y0, x1, y1], fill=self.fg_color)

        if img_size != size:
            img = img.resize((size, size), Image.LANCZOS) # type: ignore
        return img

    # ---------------------------------------------------------------------
    # Clipboard / saving helpers
    # ---------------------------------------------------------------------
    def pil_to_qimage(self, img: Image.Image) -> QImage:
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA")
        data = img.tobytes("raw", img.mode)
        if img.mode == "RGBA":
            qimg = QImage(
                data,
                img.width,
                img.height,
                img.width * 4,
                QImage.Format_RGBA8888,
            )
        else:
            qimg = QImage(
                data,
                img.width,
                img.height,
                img.width * 3,
                QImage.Format_RGB888,
            )
        return qimg.copy()

    def show_qr(self, img: Image.Image) -> None:
        qimg = self.pil_to_qimage(img)
        pix = QPixmap.fromImage(qimg)
        self.qr_label.setPixmap(pix.scaled(self.qr_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)) # type: ignore

    def copy_to_clipboard(self) -> None:
        if self.current_qr is None:
            QMessageBox.information(self, "Nothing to copy", "Generate a QR first.")
            return
        qimg = self.pil_to_qimage(self.current_qr)
        QApplication.clipboard().setImage(qimg) # type: ignore
        QMessageBox.information(self, "Copied", "QR code copied to clipboard.")

    def save_png(self) -> None:
        if self.current_qr is None:
            QMessageBox.information(self, "Nothing to save", "Generate a QR first.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save QR", "contact_qr.png", "PNG Images (*.png)")
        if path:
            self.current_qr.save(path)
            QMessageBox.information(self, "Saved", f"QR saved to {path}")


# -------------------------------------------------------------------------
# Entrypoint
# -------------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VCardQRApp()
    window.resize(900, 600)
    window.show()
    sys.exit(app.exec_())
