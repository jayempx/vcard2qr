# Changelog

## 2.0.0 — local modernization

### Fixed

- Required vCard name fields, escaped text, UTF-8 line folding and telephone type syntax.
- Double QR borders, interpolated pixels and invalid output at small dimensions.
- Batch imports replacing the image behind an unrelated single-contact preview.
- Silent batch overwrites, duplicate header collisions and unhandled row errors.
- UI freezes during Excel work and unsafe worker shutdown.
- Stale copy/save actions after changing contact details or appearance.

### Changed

- Migrated PyQt5 to PySide6 / Qt 6 with only the Essentials runtime.
- Reorganized the application into small UI, vCard, QR and Excel modules.
- Added light/dark themes, grouped fields, removable custom fields and inline validation.
- Added workbook preview, explicit column mapping, progress, cancellation and row reports.
- Added pinned dependencies, automated regression/QR-decode checks, packaging scripts and CI definitions.
- Expanded setup, privacy, Excel, architecture and validation documentation.

### Compatibility

- Python 3.12+ replaces the old Python 3.8 baseline.
- The original `python vcard-2-qr.py` launch command remains available.
- vCard 3.0 and the supplied Excel sample remain supported.
- Inputs with only an address/title/custom field now need a usable contact identity.
- Formulas are rejected instead of importing potentially stale cached results.
- Tiny or low-contrast QR settings now produce an explanatory error.
- Exports use a new run folder and row-prefixed filenames.

This is a local development version; no public release is implied by the version number.
