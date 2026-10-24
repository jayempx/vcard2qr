"""Render crisp QR modules and save exports without partial replacements."""

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path

import qrcode
from PIL import Image, ImageColor, ImageDraw
from qrcode.exceptions import DataOverflowError


@dataclass(frozen=True)
class QRSettings:
    """Immutable settings can safely be passed to a background worker."""

    size: int = 512
    radius: int = 0
    foreground: str = "#102d35"
    background: str = "#ffffff"
    transparent: bool = False


def luminance(color: str) -> float:
    """Calculate relative sRGB luminance for conservative contrast checks."""
    channels = [channel / 255 for channel in ImageColor.getrgb(color)]
    linear = [v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4 for v in channels]
    return sum(v * weight for v, weight in zip(linear, (0.2126, 0.7152, 0.0722)))


def render_qr(payload: str, settings: QRSettings = QRSettings()) -> Image.Image:
    """Preserve integer module sizes, one quiet zone and square finder patterns."""
    if not 128 <= settings.size <= 4096 or not 0 <= settings.radius <= 20:
        raise ValueError("Choose a size from 128 to 4096 px and a radius from 0 to 20 px.")
    foreground = luminance(settings.foreground)
    background = luminance("#ffffff" if settings.transparent else settings.background)
    if foreground >= background or (background + 0.05) / (foreground + 0.05) < 4.5:
        raise ValueError("Use dark modules on a light background with at least 4.5:1 contrast.")
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_Q, border=4)
    qr.add_data(payload)
    try:
        qr.make(fit=True)
    except (DataOverflowError, ValueError) as exc:
        # qrcode 8.2 may report oversized data as an invalid version 41 instead.
        raise ValueError("This contact exceeds QR capacity. Shorten its fields.") from exc
    matrix = qr.get_matrix()
    count = len(matrix)
    box = settings.size // count
    if box < 3:
        raise ValueError(f"This contact needs at least {count * 3} px. Increase the image size.")

    # Center exact-size modules inside the requested canvas; never interpolate pixels.
    offset = (settings.size - count * box) // 2
    mode = "RGBA" if settings.transparent else "RGB"
    background_color = (0, 0, 0, 0) if settings.transparent else settings.background
    image = Image.new(mode, (settings.size, settings.size), background_color)
    draw = ImageDraw.Draw(image)
    radius = min(settings.radius, box // 4)
    for row, cells in enumerate(matrix):
        for col, dark in enumerate(cells):
            if not dark:
                continue
            x, y = offset + col * box, offset + row * box
            bounds = (x, y, x + box - 1, y + box - 1)
            finder = (
                (row < 12 and col < 12)
                or (row < 12 and col >= count - 12)
                or (row >= count - 12 and col < 12)
            )
            draw.rounded_rectangle(bounds, radius=0 if finder else radius, fill=settings.foreground)
    return image


def save_png(image: Image.Image, path: str | Path) -> None:
    """Replace an approved destination only after encoding has succeeded."""
    path = Path(path)
    descriptor, temporary = tempfile.mkstemp(prefix=".vcard2qr-", suffix=".png", dir=path.parent)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            image.save(stream, format="PNG")
        os.replace(temporary, path)
    finally:
        # Remove only our own temporary file after a failed encode or replacement.
        Path(temporary).unlink(missing_ok=True)
