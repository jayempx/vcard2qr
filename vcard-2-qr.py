"""Compatibility launcher for existing shortcuts and Windows packaging."""

from vcard2qr.app import main

# Keep the original command working after moving the implementation into a package.
if __name__ == "__main__":
    raise SystemExit(main())
