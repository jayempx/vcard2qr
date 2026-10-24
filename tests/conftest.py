"""Run native Qt widgets without opening windows during automated tests."""

import os

import pytest

# Set the platform before Qt creates its application instance.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture(scope="session")
def qapp():
    from PySide6.QtWidgets import QApplication

    application = QApplication.instance() or QApplication([])
    application.setStyle("Fusion")
    return application
