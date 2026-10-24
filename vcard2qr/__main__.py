"""Allow python -m vcard2qr as well as the installed desktop entry point."""

from .app import main

if __name__ == "__main__":
    raise SystemExit(main())
