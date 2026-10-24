"""PySide6 desktop UI; contact snapshots and batch jobs never share output state."""

import argparse
import sys
from collections.abc import Callable
from pathlib import Path

from PIL import Image
from PySide6.QtCore import QRect, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QColor, QImage, QPainter, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QColorDialog,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from .excel import default_mapping, export_workbook, inspect_workbook, validate_mapping
from .qr import QRSettings, render_qr, save_png
from .vcard import FIELDS, ContactError, build_vcard


def qt_image(image: Image.Image) -> QImage:
    """Copy the Pillow buffer so Qt owns its image after this function returns."""
    rgba = image.convert("RGBA")
    return QImage(
        rgba.tobytes(), rgba.width, rgba.height, rgba.width * 4, QImage.Format.Format_RGBA8888
    ).copy()


class Preview(QWidget):
    """Paint transparency over a checkerboard without modifying the exported image."""

    def __init__(self):
        super().__init__()
        self.image = None
        self.setMinimumSize(220, 220)
        self.setAccessibleName("QR preview")

    def paintEvent(self, event):
        painter = QPainter(self)
        side = min(self.width(), self.height()) - 16
        target = QRect((self.width() - side) // 2, (self.height() - side) // 2, side, side)
        for y in range(target.top(), target.bottom() + 1, 12):
            for x in range(target.left(), target.right() + 1, 12):
                shade = "#ffffff" if ((x - target.left()) // 12 + (y - target.top()) // 12) % 2 else "#e8edef"
                painter.fillRect(QRect(x, y, 12, 12).intersected(target), QColor(shade))
        if self.image is not None:
            painter.drawImage(target, self.image)
        else:
            painter.setPen(QColor("#53676d"))
            painter.drawText(
                target,
                Qt.AlignmentFlag.AlignCenter,
                "Your next connection\nstarts here.\n\nCreate a contact to preview its QR.",
            )
        painter.end()


class Task(QThread):
    """Execute file work off the UI thread, using only immutable input snapshots."""

    succeeded = Signal(object)
    failed = Signal(str)
    progress = Signal(int, int)

    def __init__(self, kind: str, operation: Callable, parent=None):
        super().__init__(parent)
        self.kind, self.operation = kind, operation

    def run(self):
        try:
            self.succeeded.emit(self.operation(self))
        except Exception as exc:
            # Keep third-party workbook and filesystem errors inside the job boundary.
            self.failed.emit(f"{type(exc).__name__}: {exc}")


class VCardQRApp(QWidget):
    """A two-mode editor with explicit generation and safe export controls."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("vcard2qr · Contact studio")
        self.resize(1100, 850)
        self.setMinimumSize(880, 650)
        self.current_qr = None
        self.current_payload = None
        self.worker = None
        self.closing = False
        self.book = None
        self.book_path = ""
        self.custom_fields = []
        self.fields, self.errors = {}, {}
        self.fg_color, self.bg_color = "#102d35", "#ffffff"

        # Header establishes purpose, privacy and a small theme control.
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(18)
        header = QHBoxLayout()
        heading = QVBoxLayout()
        title = QLabel("vcard2qr")
        title.setObjectName("brand")
        heading.addWidget(title)
        heading.addWidget(QLabel("Good connections. One scan away."))
        header.addLayout(heading)
        header.addStretch()
        badge = QLabel("LOCAL ONLY  ·  NO UPLOADS")
        badge.setObjectName("badge")
        header.addWidget(badge)
        self.theme = QComboBox()
        self.theme.addItems(["Light", "Dark"])
        self.theme.setAccessibleName("Color theme")
        self.theme.currentTextChanged.connect(self.apply_theme)
        header.addWidget(self.theme)
        root.addLayout(header)

        # The splitter leaves the form scrollable while preserving a useful preview.
        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.tabs = QTabWidget()
        self.tabs.addTab(self.contact_tab(), "&Single contact")
        self.tabs.addTab(self.batch_tab(), "&Excel import")
        splitter.addWidget(self.tabs)
        splitter.addWidget(self.output_panel())
        splitter.setSizes([580, 420])
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)
        self.status = QLabel("Ready. Your contact data stays on this computer.")
        self.status.setWordWrap(True)
        self.status.setAccessibleName("Application status")
        root.addWidget(self.status)
        self.apply_theme("Light")
        self.invalidate()

    def contact_tab(self):
        """Group fields by purpose and keep validation messages beside each input."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(16, 20, 16, 16)
        for title, keys in (
            ("01  Identity", FIELDS[:4]),
            ("02  Contact details", FIELDS[4:]),
        ):
            group = QGroupBox(title)
            form = QFormLayout(group)
            form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
            form.setVerticalSpacing(12)
            for key in keys:
                box = QWidget()
                column = QVBoxLayout(box)
                column.setContentsMargins(0, 0, 0, 0)
                column.setSpacing(2)
                edit, error = QLineEdit(), QLabel()
                edit.setAccessibleName(key)
                edit.setPlaceholderText(
                    {
                        "Email": "name@example.com",
                        "Mobile": "+39 333 1234567",
                        "LinkedIn": "https://linkedin.com/in/your-name",
                    }.get(key, key)
                )
                edit.textChanged.connect(self.invalidate)
                error.setObjectName("error")
                error.setWordWrap(True)
                error.hide()
                column.addWidget(edit)
                column.addWidget(error)
                label = QLabel(key)
                label.setBuddy(edit)
                form.addRow(label, box)
                self.fields[key], self.errors[key] = edit, error
            layout.addWidget(group)
        custom = QGroupBox("03  Custom fields")
        self.custom_layout = QVBoxLayout(custom)
        add = QPushButton("+ Add field")
        add.clicked.connect(self.add_custom_field)
        self.custom_layout.addWidget(add)
        layout.addWidget(custom)
        layout.addStretch()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(page)
        return scroll

    def add_custom_field(self):
        """Custom fields can be removed as well as added."""
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        key, value = QLineEdit(), QLineEdit()
        key.setPlaceholderText("Field name")
        value.setPlaceholderText("Value")
        key.setAccessibleName("Custom field name")
        value.setAccessibleName("Custom field value")
        remove = QPushButton("Remove")
        for widget in (key, value, remove):
            layout.addWidget(widget)
        entry = (key, value)
        self.custom_fields.append(entry)
        key.textChanged.connect(self.invalidate)
        value.textChanged.connect(self.invalidate)

        def remove_row():
            self.custom_fields.remove(entry)
            row.deleteLater()
            self.invalidate()

        remove.clicked.connect(remove_row)
        self.custom_layout.addWidget(row)
        self.invalidate()

    def batch_tab(self):
        """Expose header mapping and a small sample before any export is started."""
        page = QWidget()
        layout = QVBoxLayout(page)
        intro = QLabel("Import contacts in one pass. Review the columns, then export to a new folder.")
        intro.setWordWrap(True)
        layout.addWidget(intro)
        self.load_button = QPushButton("Choose Excel workbook…")
        self.load_button.clicked.connect(self.load_excel)
        layout.addWidget(self.load_button)
        self.book_label = QLabel("No workbook selected · .xlsx files")
        self.book_label.setWordWrap(True)
        layout.addWidget(self.book_label)
        self.mapping_table = QTableWidget(0, 3)
        self.mapping_table.setHorizontalHeaderLabels(["Source column", "Use as", "Example"])
        self.mapping_table.verticalHeader().setDefaultSectionSize(42)
        self.mapping_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.mapping_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.mapping_table.setAccessibleName("Excel column mapping")
        layout.addWidget(self.mapping_table, 2)
        layout.addWidget(QLabel("First five rows · preview only"))
        self.sample_table = QTableWidget()
        self.sample_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.sample_table.setAccessibleName("Workbook sample rows")
        self.sample_table.setMaximumHeight(180)
        layout.addWidget(self.sample_table, 1)
        note = QLabel(
            "Phone numbers should be text in Excel to preserve + and leading zeros.\n"
            "Formulas are rejected. Existing exports are never overwritten."
        )
        note.setWordWrap(True)
        layout.addWidget(note)
        buttons = QHBoxLayout()
        self.batch_button = QPushButton("Export contacts…")
        self.batch_button.setObjectName("primary")
        self.batch_button.setEnabled(False)
        self.batch_button.clicked.connect(self.start_batch)
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel_job)
        buttons.addWidget(self.batch_button)
        buttons.addWidget(self.cancel_button)
        layout.addLayout(buttons)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)
        self.batch_summary = QLabel("Uses the appearance settings on the right.")
        self.batch_summary.setWordWrap(True)
        self.batch_summary.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(self.batch_summary)
        return page

    def output_panel(self):
        """Keep shared rendering settings separate from the single-contact snapshot."""
        panel = QWidget()
        panel.setMinimumWidth(320)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(18, 8, 0, 0)
        caption = QLabel("SINGLE CONTACT PREVIEW")
        caption.setObjectName("eyebrow")
        layout.addWidget(caption)
        self.preview = Preview()
        layout.addWidget(self.preview, 1)
        self.preview_status = QLabel("No QR generated yet")
        self.preview_status.setWordWrap(True)
        layout.addWidget(self.preview_status)
        appearance = QGroupBox("Appearance · both modes")
        form = QFormLayout(appearance)
        self.size_spin = QSpinBox()
        self.size_spin.setRange(128, 4096)
        self.size_spin.setValue(512)
        self.size_spin.setSuffix(" px")
        self.size_spin.setAccessibleName("Output size")
        self.radius_spin = QSpinBox()
        self.radius_spin.setRange(0, 20)
        self.radius_spin.setSuffix(" px")
        self.radius_spin.setAccessibleName("Module corner radius")
        form.addRow("Image size", self.size_spin)
        form.addRow("Corner radius", self.radius_spin)
        self.fg_button = QPushButton()
        self.bg_button = QPushButton()
        self.fg_button.clicked.connect(lambda: self.pick_color(True))
        self.bg_button.clicked.connect(lambda: self.pick_color(False))
        form.addRow("Foreground", self.fg_button)
        form.addRow("Background", self.bg_button)
        self.transparent = QCheckBox("Transparent background")
        self.transparent.toggled.connect(self.change_transparency)
        form.addRow(self.transparent)
        layout.addWidget(appearance)
        hint = QLabel("Transparency needs a light surface. Always scan-test before printing.")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        self.generate_button = QPushButton("&Generate QR")
        self.generate_button.setObjectName("primary")
        self.generate_button.setShortcut("Ctrl+Return")
        self.generate_button.clicked.connect(self.generate_qr)
        layout.addWidget(self.generate_button)
        actions = QHBoxLayout()
        self.copy_button = QPushButton("&Copy image")
        self.save_button = QPushButton("Save &PNG…")
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        self.save_button.clicked.connect(self.save_current)
        actions.addWidget(self.copy_button)
        actions.addWidget(self.save_button)
        layout.addLayout(actions)
        self.size_spin.valueChanged.connect(self.invalidate)
        self.radius_spin.valueChanged.connect(self.invalidate)
        self.update_color_labels()
        # Keep export actions reachable on short screens and at high DPI.
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidget(panel)
        scroll.setMinimumWidth(340)
        return scroll

    def settings(self) -> QRSettings:
        return QRSettings(
            self.size_spin.value(),
            self.radius_spin.value(),
            self.fg_color,
            self.bg_color,
            self.transparent.isChecked(),
        )

    def update_color_labels(self):
        """Hex labels and color chips make selected colors visible in either theme."""
        for button, color in ((self.fg_button, self.fg_color), (self.bg_button, self.bg_color)):
            button.setText(color.upper())
            button.setStyleSheet(f"border-left: 16px solid {color};")

    def pick_color(self, foreground: bool):
        initial = self.fg_color if foreground else self.bg_color
        color = QColorDialog.getColor(QColor(initial), self, "Choose QR color")
        if color.isValid():
            if foreground:
                self.fg_color = color.name()
            else:
                self.bg_color = color.name()
            self.update_color_labels()
            self.invalidate()

    def change_transparency(self, checked: bool):
        self.bg_button.setEnabled(not checked)
        self.invalidate()

    def invalidate(self, *_):
        """A changed input immediately disables export and removes the old image."""
        if not hasattr(self, "preview_status"):
            return
        had_output = self.current_qr is not None
        self.current_qr = self.current_payload = None
        self.preview.image = None
        self.preview.update()
        self.copy_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.preview_status.setText("Inputs changed · generate again" if had_output else "Ready to generate")
        for error in self.errors.values():
            error.hide()

    def generate_qr(self):
        """Publish a new snapshot only when validation and rendering both succeed."""
        self.invalidate()
        try:
            payload = build_vcard(
                {key: edit.text() for key, edit in self.fields.items()},
                [(key.text(), value.text()) for key, value in self.custom_fields],
            )
            image = render_qr(payload, self.settings())
        except (ValueError, OSError) as exc:
            self.status.setText(str(exc))
            if isinstance(exc, ContactError) and exc.field in self.errors:
                self.errors[exc.field].setText(str(exc))
                self.errors[exc.field].show()
                self.tabs.setCurrentIndex(0)
                self.fields[exc.field].setFocus()
            return
        self.current_qr, self.current_payload = image, payload
        self.preview.image = qt_image(image)
        self.preview.update()
        self.preview_status.setText(f"Ready to share · {image.width} × {image.height} px")
        self.copy_button.setEnabled(True)
        self.save_button.setEnabled(True)
        self.status.setText("QR generated. Scan-test it with the contact app you plan to use.")

    def copy_to_clipboard(self):
        if self.current_qr is not None:
            clipboard = QApplication.clipboard()
            image = qt_image(self.current_qr)
            clipboard.setImage(image)
            # Some restricted Windows sessions deny clipboard access without an exception.
            if clipboard.ownsClipboard() or clipboard.image() == image:
                self.status.setText("Image copied to the clipboard.")
            else:
                self.status.setText("Clipboard unavailable. Try again or use Save PNG instead.")

    def save_current(self):
        """The native dialog confirms replacement; atomic save protects old files."""
        if self.current_qr is None:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save QR image", "contact_qr.png", "PNG images (*.png)")
        if not path:
            return
        try:
            save_png(self.current_qr, path)
            self.status.setText(f"Saved {path}")
        except OSError as exc:
            self.status.setText(f"Could not save the image: {exc}")

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose contacts", "", "Excel workbooks (*.xlsx)")
        if path:
            self.load_workbook_path(path)

    def load_workbook_path(self, path: str):
        """Load a preview asynchronously; replacing a workbook clears old mapping."""
        if self.worker is not None:
            return
        self.book, self.book_path = None, path
        self.mapping_table.setRowCount(0)
        self.sample_table.setRowCount(0)
        self.book_label.setText("Reading workbook…")
        self.begin_job("inspect", lambda task: inspect_workbook(path))

    def begin_job(self, kind: str, operation: Callable):
        self.worker = Task(kind, operation, self)
        self.worker.succeeded.connect(self.job_result)
        self.worker.failed.connect(self.job_error)
        self.worker.progress.connect(self.job_progress)
        self.worker.finished.connect(self.job_finished)
        self.load_button.setEnabled(False)
        self.batch_button.setEnabled(False)
        self.mapping_table.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.progress.setRange(0, 0)
        self.batch_summary.setText("Reading workbook…" if kind == "inspect" else "Exporting…")
        self.worker.start()

    def job_progress(self, done: int, total: int):
        self.progress.setRange(0, max(done, total, 1))
        self.progress.setValue(done)

    def job_result(self, result: dict):
        """Batch results never overwrite the single-contact preview or clipboard source."""
        if self.worker.kind == "inspect":
            if self.worker.isInterruptionRequested():
                self.book_label.setText("Workbook loading cancelled")
                return
            self.book = result
            headers, preview = result["headers"], result["preview"]
            self.book_label.setText(
                f"{Path(self.book_path).name} · {result['sheet']} · about {result['rows']} rows"
            )
            self.mapping_table.setRowCount(len(headers))
            for index, (header, suggested) in enumerate(zip(headers, default_mapping(headers))):
                self.mapping_table.setItem(index, 0, QTableWidgetItem(header or "(blank)"))
                combo = QComboBox()
                combo.addItems(["Ignore", *FIELDS, f"Custom: {header}"])
                combo.setCurrentText(suggested)
                combo.setAccessibleName(f"Mapping for column {index + 1}: {header}")
                self.mapping_table.setCellWidget(index, 1, combo)
                sample = preview[0][index] if preview and index < len(preview[0]) else ""
                self.mapping_table.setItem(index, 2, QTableWidgetItem(sample))
            self.sample_table.setColumnCount(len(headers))
            self.sample_table.setHorizontalHeaderLabels(headers)
            self.sample_table.setRowCount(len(preview))
            for row, values in enumerate(preview):
                for col, value in enumerate(values):
                    self.sample_table.setItem(row, col, QTableWidgetItem(value))
            self.batch_summary.setText(
                "Review the mappings. Unknown columns are custom fields; use Ignore to exclude them."
            )
        else:
            state = "Cancelled" if result["cancelled"] else "Complete"
            self.batch_summary.setText(
                f"{state} · {result['saved']} saved · {result['skipped']} skipped · "
                f"{result['failed']} failed\n{result['folder']}\nSee report.jsonl for row details."
            )
            self.status.setText("Batch finished. The single-contact preview is unchanged.")

    def job_error(self, error: str):
        self.batch_summary.setText(f"Unable to complete the operation: {error}")
        if self.worker.kind == "inspect":
            self.book_label.setText("Workbook could not be loaded")

    def job_finished(self):
        """Release the thread only after it has stopped; never terminate file writes."""
        self.worker.deleteLater()
        self.worker = None
        self.load_button.setEnabled(True)
        self.batch_button.setEnabled(self.book is not None)
        self.mapping_table.setEnabled(True)
        self.cancel_button.setEnabled(False)
        if self.progress.maximum() == 0:
            self.progress.setRange(0, 1)
            self.progress.setValue(0)
        if self.closing:
            QTimer.singleShot(0, self.close)

    def start_batch(self):
        if self.book is None or self.worker is not None:
            return
        mapping = [
            self.mapping_table.cellWidget(index, 1).currentText()
            for index in range(self.mapping_table.rowCount())
        ]
        try:
            validate_mapping(mapping)
        except ValueError as exc:
            self.batch_summary.setText(str(exc))
            return
        output = QFileDialog.getExistingDirectory(self, "Choose parent folder for a new export")
        if output:
            self.export_to(output, mapping)

    def export_to(self, output: str, mapping: list[str]):
        """Capture all settings before the worker starts; UI edits affect the next run."""
        if self.worker is not None or self.book is None:
            return
        path, headers, settings = self.book_path, tuple(self.book["headers"]), self.settings()
        targets = tuple(mapping)
        self.begin_job(
            "export",
            lambda task: export_workbook(
                path, output, headers, targets, settings, task.isInterruptionRequested, task.progress.emit
            ),
        )

    def cancel_job(self):
        if self.worker is not None:
            self.worker.requestInterruption()
            self.cancel_button.setEnabled(False)
            self.batch_summary.setText("Cancelling after the current row. Completed files will be kept.")

    def closeEvent(self, event):
        if self.worker is not None:
            self.closing = True
            self.cancel_job()
            event.ignore()
        else:
            event.accept()

    def apply_theme(self, theme: str):
        """Use supported Qt styles only, with visible focus and disabled states."""
        dark = theme == "Dark"
        bg, panel, text, muted, border = (
            ("#112127", "#192f36", "#edf6f5", "#afc5ca", "#3e5961")
            if dark
            else ("#f4f7f6", "#ffffff", "#17333c", "#536b73", "#cbdad9")
        )
        # Native arrows, checkboxes and popup menus also need the matching palette.
        palette = self.palette()
        for role, color in (
            (QPalette.ColorRole.Window, bg),
            (QPalette.ColorRole.Base, panel),
            (QPalette.ColorRole.Text, text),
            (QPalette.ColorRole.WindowText, text),
            (QPalette.ColorRole.Button, panel),
            (QPalette.ColorRole.ButtonText, text),
            (QPalette.ColorRole.Highlight, "#087e72"),
            (QPalette.ColorRole.HighlightedText, "#ffffff"),
        ):
            palette.setColor(role, QColor(color))
        self.setPalette(palette)
        self.setStyleSheet(f"""
            /* Base surfaces and typography. */
            QWidget {{ background: {bg}; color: {text}; font-family: 'Segoe UI'; font-size: 13px; }}
            QLabel {{ background: transparent; }}
            QLabel#brand {{ font-size: 30px; font-weight: 700; letter-spacing: -1px; }}
            QLabel#badge {{ color: {muted}; font-size: 11px; padding: 8px; }}
            QLabel#eyebrow {{ color: {muted}; font-weight: 600; font-size: 11px; }}
            QLabel#error {{ color: {"#ffaaa3" if dark else "#b32e26"}; font-size: 12px; }}
            QGroupBox {{ border: 1px solid {border}; border-radius: 10px; margin-top: 14px; padding: 18px 12px 10px; }}
            QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 5px; font-weight: 600; }}
            /* Inputs keep native behavior and explicit keyboard focus. */
            QLineEdit, QSpinBox, QComboBox {{ background: {panel}; border: 1px solid {border}; border-radius: 6px; padding: 7px; min-height: 20px; }}
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus, QPushButton:focus {{ border: 2px solid #10a797; }}
            QPushButton {{ background: {panel}; border: 1px solid {border}; border-radius: 7px; padding: 9px 12px; min-height: 20px; font-weight: 600; }}
            QPushButton:hover {{ border-color: #108578; }}
            QPushButton#primary {{ background: #087e72; color: white; border-color: #087e72; }}
            QPushButton#primary:hover {{ background: #06695f; }}
            QPushButton:disabled {{ color: {muted}; background: {bg}; }}
            QTabWidget::pane {{ border: 1px solid {border}; border-radius: 8px; }}
            QTabBar::tab {{ padding: 12px 20px; color: {muted}; }}
            QTabBar::tab:selected {{ color: {text}; border-bottom: 3px solid #0a9385; }}
            QTableWidget {{ background: {panel}; gridline-color: {border}; border: 1px solid {border}; selection-background-color: #087e72; selection-color: white; }}
            QHeaderView::section {{ background: {bg}; color: {text}; padding: 6px; border: 0; }}
            QProgressBar {{ border: 1px solid {border}; border-radius: 5px; text-align: center; min-height: 18px; }}
            QProgressBar::chunk {{ background: #0a9385; }}
            QScrollArea {{ border: 0; }}
        """)


def main() -> int:
    """Start the native application without network services or contact persistence."""
    parser = argparse.ArgumentParser(description="Local contact QR studio")
    parser.add_argument(
        "--smoke-test", type=Path, metavar="OUTPUT", help="Verify a packaged build with synthetic contacts"
    )
    args = parser.parse_args()
    app = QApplication(sys.argv)
    app.setApplicationName("vcard2qr")
    app.setStyle("Fusion")
    window = VCardQRApp()
    if args.smoke_test:
        from .smoke import run

        return run(app, window, args.smoke_test)
    window.show()
    return app.exec()
