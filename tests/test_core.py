"""Regression coverage for real payloads, raster decoding and batch boundaries."""

import json
from dataclasses import replace

import pytest
import zxingcpp
from openpyxl import Workbook
from PIL import Image

from vcard2qr.excel import default_mapping, export_workbook, inspect_workbook, validate_mapping
from vcard2qr.qr import QRSettings, render_qr, save_png
from vcard2qr.vcard import ContactError, build_vcard


def make_book(path, rows):
    """Create synthetic contacts only; no real address books enter the tests."""
    book = Workbook()
    for row in rows:
        book.active.append(row)
    book.save(path)
    book.close()


def test_vcard_escaping_required_names_and_phone_types():
    payload = build_vcard(
        {
            "First Name": "Zoë",
            "Last Name": "Rossi",
            "Organization": "Studio;Partner\nTITLE:Injected",
            "Direct Office": "+39 051 123456",
            "Address": "Via A, 2\\B",
        },
        [("Customer ID", "A;B,C")],
    )
    assert "N:Rossi;Zoë;;;\r\n" in payload
    assert "FN:Zoë Rossi\r\n" in payload
    assert "ORG:Studio\\;Partner\\nTITLE:Injected\r\n" in payload
    assert "\r\nTITLE:Injected" not in payload
    assert "TEL;TYPE=WORK,VOICE:+39 051 123456\r\n" in payload
    assert "ADR;TYPE=WORK:;;Via A\\, 2\\\\B;;;;\r\n" in payload
    assert "X-CUSTOMER-ID:A\\;B\\,C\r\n" in payload
    assert payload.endswith("END:VCARD\r\n")
    assert "\n" not in payload.replace("\r\n", "")


def test_unicode_folding_preserves_octets():
    name = "東京É🙂" * 60
    payload = build_vcard({"First Name": name})
    assert all(len(line.encode("utf-8")) <= 75 for line in payload.split("\r\n"))
    assert f"FN:{name}\r\n" in payload.replace("\r\n ", "")


@pytest.mark.parametrize(
    "data,custom",
    [
        ({}, []),
        ({"Title": "No identity"}, []),
        ({"Email": "broken"}, []),
        ({"First Name": "Ada", "LinkedIn": "javascript:alert(1)"}, []),
        ({"Mobile": "123\nFN:Injected"}, []),
        ({"First Name": "Ada"}, [("Name:Injected", "value")]),
        ({"First Name": "Ada"}, [("a b", "1"), ("a-b", "2")]),
        ({"First Name": "Ada"}, [("Only name", "")]),
    ],
)
def test_invalid_contacts_fail_before_rendering(data, custom):
    with pytest.raises(ContactError):
        build_vcard(data, custom)


def test_organization_only_has_required_name_properties():
    payload = build_vcard({"Organization": "Example"})
    assert "N:;;;;\r\nFN:Example\r\n" in payload


@pytest.mark.parametrize(
    "transparent,radius,size", [(False, 0, 512), (False, 20, 1024), (True, 0, 512), (True, 20, 1024)]
)
def test_exported_png_decodes_to_exact_payload(tmp_path, transparent, radius, size):
    payload = build_vcard(
        {"First Name": "Zoë", "Last Name": "Rossi", "Mobile": "+39 333 0123456", "Email": "zoe@example.com"}
    )
    image = render_qr(payload, QRSettings(size=size, transparent=transparent, radius=radius))
    path = tmp_path / "contact.png"
    save_png(image, path)
    with Image.open(path) as exported:
        assert exported.size == (size, size)
        if transparent:
            assert exported.getpixel((0, 0))[3] == 0
            surface = Image.new("RGBA", exported.size, "white")
            surface.alpha_composite(exported)
            exported = surface.convert("RGB")
        decoded = zxingcpp.read_barcode(exported)
        assert decoded is not None
        assert decoded.bytes == payload.encode("utf-8")


def test_raster_has_no_interpolation_and_only_one_quiet_zone():
    image = render_qr(build_vcard({"First Name": "Ada"}), QRSettings(foreground="#000000"))
    assert len(image.getcolors()) == 2
    # The first dark module starts after four modules plus centered leftover pixels.
    import qrcode

    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_Q, border=4)
    qr.add_data(build_vcard({"First Name": "Ada"}))
    qr.make(fit=True)
    count = len(qr.get_matrix())
    box, offset = 512 // count, (512 - (512 // count) * count) // 2
    assert image.getpixel((offset + 4 * box, offset + 4 * box)) == (0, 0, 0)
    assert image.getpixel((offset + 4 * box - 1, offset + 4 * box)) == (255, 255, 255)


def test_small_dense_qr_and_overflow_are_actionable_errors():
    with pytest.raises(ValueError, match="at least"):
        render_qr(build_vcard({"First Name": "Ada", "Title": "X" * 1000}), QRSettings(size=128))
    with pytest.raises(ValueError, match="capacity"):
        render_qr("x" * 10000)


@pytest.mark.parametrize(
    "changes", [{"foreground": "#ffffff"}, {"background": "#102d35"}, {"size": 0}, {"radius": -1}]
)
def test_invalid_render_settings(changes):
    with pytest.raises(ValueError):
        render_qr("test", replace(QRSettings(), **changes))


def test_failed_atomic_save_preserves_existing_file(tmp_path, monkeypatch):
    target = tmp_path / "existing.png"
    target.write_bytes(b"original")
    image = Image.new("RGB", (10, 10))

    def fail(*args, **kwargs):
        raise OSError("disk full")

    monkeypatch.setattr(image, "save", fail)
    with pytest.raises(OSError):
        save_png(image, target)
    assert target.read_bytes() == b"original"
    assert list(tmp_path.iterdir()) == [target]


def test_batch_reports_errors_continues_and_never_overwrites(tmp_path):
    path = tmp_path / "contacts.xlsx"
    make_book(
        path,
        [
            ["Nome", "Email", "Cellulare"],
            ["Ada", "ada@example.com", "00123"],
            ["Bad", "broken", ""],
            [None, None, None],
            ["Formula", "=1+1", ""],
            ["Ada", "ada@example.com", 123],
        ],
    )
    book = inspect_workbook(path)
    mapping = default_mapping(book["headers"])
    first = export_workbook(path, tmp_path, book["headers"], mapping, QRSettings())
    second = export_workbook(path, tmp_path, book["headers"], mapping, QRSettings())
    assert first["folder"] != second["folder"]
    assert (first["saved"], first["skipped"], first["failed"]) == (2, 1, 2)
    from pathlib import Path

    folder = Path(first["folder"])
    records = [json.loads(line) for line in (folder / "report.jsonl").read_text().splitlines()]
    assert records[-1]["warnings"]
    decoded = zxingcpp.read_barcode(Image.open(folder / records[0]["file"]))
    assert b"TEL;TYPE=CELL:00123" in decoded.bytes


def test_batch_cancellation_keeps_completed_files(tmp_path):
    path = tmp_path / "contacts.xlsx"
    make_book(path, [["First Name"], ["One"], ["Two"], ["Three"]])
    processed = []
    result = export_workbook(
        path,
        tmp_path,
        ["First Name"],
        ["First Name"],
        QRSettings(),
        cancelled=lambda: bool(processed),
        progress=lambda done, total: processed.append(done),
    )
    assert result["saved"] == 1
    assert result["cancelled"] is True


def test_mapping_collisions_and_changed_workbook_are_rejected(tmp_path):
    with pytest.raises(ValueError, match="More than one"):
        validate_mapping(default_mapping(["Nome", "First Name"]))
    with pytest.raises(ValueError):
        validate_mapping(["Ignore"])
    path = tmp_path / "contacts.xlsx"
    make_book(path, [["Email"], ["a@example.com"]])
    with pytest.raises(ValueError, match="headers changed"):
        export_workbook(path, tmp_path, ["Nome"], ["First Name"], QRSettings())
    assert list(tmp_path.iterdir()) == [path]


def test_supplied_workbook_remains_compatible(tmp_path):
    from pathlib import Path

    path = Path(__file__).resolve().parents[1] / "sample_contacts.xlsx"
    book = inspect_workbook(path)
    result = export_workbook(path, tmp_path, book["headers"], default_mapping(book["headers"]), QRSettings())
    assert result["saved"] == 2
    assert result["failed"] == 0


def test_small_valid_qr_decodes():
    payload = build_vcard({"First Name": "Ada"})
    decoded = zxingcpp.read_barcode(render_qr(payload, QRSettings(size=135)))
    assert decoded is not None
    assert decoded.bytes == payload.encode("utf-8")
