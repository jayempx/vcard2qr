# vcard2qr

WARNING: this a vibe-coded repo made with Codex.

Simple PyQt5 desktop app that collects your contact details, builds a vCard 3.0 payload, and renders it as a customizable QR code you can preview, copy or export as PNG called **vcard2qr**.

## Features

- Fill name, title, email, phones, address, LinkedIn, and optional custom fields.
- Pick foreground / background colors and toggle a transparent background for the output image.
- Adjust QR size and corner radius before generation.
- Preview the generated code in the UI, copy it to the clipboard, or save it as PNG.
- Import multiple contacts from an Excel workbook and export the matching QR codes in one go.

## Requirements

- Python 3.8+

Install dependencies:
```
pip install -r requirements.txt
```

## Usage

1. Run the app:
   ```
   python vcard-2-qr.py
   ```
2. Fill at least one form field to compose the vCard payload.
3. Adjust size, radius, and colors. Check “No background” if you want a transparent background.
4. Click **Generate QR**. The preview updates with the rendered QR.
5. Use **Copy to clipboard** or **Save as PNG …** to export the image.
6. Choose **Batch convert …** to load a workbook and emit PNGs for every contact using the current color/size settings.

You can use the "Batch convert..." button to import a .xlsx database file of multiple contacts and export their qr all at once, keeping valid all the settings above.

## Notes

- The PNG output respects the background selection (color or transparency).
- Clipboard copy converts the Pillow image to a Qt image before copying, so transparency is preserved.