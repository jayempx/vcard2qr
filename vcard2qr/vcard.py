"""Validate contact data and serialize interoperable UTF-8 vCard 3.0 records."""

import re
from collections.abc import Mapping, Sequence
from urllib.parse import urlsplit

# Stable field keys are shared by the form and spreadsheet mapping.
FIELDS = (
    "First Name",
    "Last Name",
    "Organization",
    "Title",
    "Email",
    "Mobile",
    "Switchboard",
    "Direct Office",
    "Address",
    "LinkedIn",
)
PHONE_TYPES = {"Mobile": "CELL", "Switchboard": "WORK", "Direct Office": "WORK,VOICE"}


class ContactError(ValueError):
    """A user-correctable validation error associated with one form field."""

    def __init__(self, field: str, message: str):
        super().__init__(message)
        self.field = field


def escape_text(value: str) -> str:
    """Escape delimiters and newlines without allowing injected properties."""
    return (
        value.replace("\\", "\\\\")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .replace("\n", "\\n")
        .replace(";", "\\;")
        .replace(",", "\\,")
    )


def fold_line(line: str) -> str:
    """Fold at 75 UTF-8 octets, keeping multibyte characters intact."""
    parts, current, size = [], "", 0
    for char in line:
        width = len(char.encode("utf-8"))
        if size + width > 75:
            parts.append(current)
            current, size = " ", 1
        current += char
        size += width
    return "\r\n".join([*parts, current])


def custom_name(key: str) -> str:
    """Accept readable extension names; reject vCard syntax in field names."""
    label = re.sub(r"[ _]+", "-", key.strip().upper())
    if label.startswith("X-"):
        label = label[2:]
    if not re.fullmatch(r"[A-Z0-9]+(?:-[A-Z0-9]+)*", label):
        raise ContactError("Custom fields", "Custom names must use letters A-Z, digits, spaces or hyphens.")
    return "X-" + label


def build_vcard(data: Mapping[str, str], custom: Sequence[tuple[str, str]] = ()) -> str:
    """Build one record; require a usable identity and validate typed fields."""
    values = {key: str(data.get(key, "")).strip() for key in FIELDS}
    for key, value in [*values.items(), *custom]:
        if any(ord(char) < 32 and char not in "\r\n\t" for char in value):
            raise ContactError(key, "Remove unsupported control characters.")
    first, last = values["First Name"], values["Last Name"]
    name = " ".join(filter(None, (first, last)))
    name = name or next((values[key] for key in ("Organization", "Email", *PHONE_TYPES) if values[key]), "")
    if not name:
        raise ContactError("First Name", "Enter a name, organization, email or phone number.")

    # URI and telephone values must never contain line breaks or property separators.
    email = values["Email"]
    if email and not re.fullmatch(r"[^\s@,;:]+@[^\s@,;:]+\.[^\s@,;:]+", email):
        raise ContactError("Email", "Enter an email address such as name@example.com.")
    for key in PHONE_TYPES:
        phone = values[key]
        if phone and (not re.fullmatch(r"[+0-9(). #xXextEXT-]+", phone) or not re.search(r"\d", phone)):
            raise ContactError(key, "Use digits, +, spaces, parentheses or an extension (ext. 123).")
    url = values["LinkedIn"]
    if url:
        try:
            parsed = urlsplit(url)
            valid_url = parsed.scheme in ("http", "https") and bool(parsed.hostname)
        except ValueError:
            valid_url = False
        if not valid_url or any(char.isspace() or ord(char) < 32 for char in url):
            raise ContactError("LinkedIn", "Enter a complete URL beginning with https:// or http://.")

    # N and FN are required by vCard 3.0, including organization-only contacts.
    lines = [
        "BEGIN:VCARD",
        "VERSION:3.0",
        f"N:{escape_text(last)};{escape_text(first)};;;",
        f"FN:{escape_text(name)}",
    ]
    for key, prop in (("Organization", "ORG"), ("Title", "TITLE")):
        if values[key]:
            lines.append(f"{prop}:{escape_text(values[key])}")
    for key, kind in PHONE_TYPES.items():
        if values[key]:
            lines.append(f"TEL;TYPE={kind}:{values[key]}")
    if email:
        lines.append(f"EMAIL;TYPE=INTERNET:{escape_text(email)}")
    if values["Address"]:
        lines.append(f"ADR;TYPE=WORK:;;{escape_text(values['Address'])};;;;")
    if url:
        lines.append(f"URL:{url}")
    names = set()
    for key, value in custom:
        key, value = key.strip(), value.strip()
        if not key and not value:
            continue
        if not key or not value:
            raise ContactError("Custom fields", "Each custom field needs both a name and a value.")
        label = custom_name(key)
        if label in names:
            raise ContactError("Custom fields", "Custom field names must be unique.")
        names.add(label)
        lines.append(f"{label}:{escape_text(value)}")
    lines.append("END:VCARD")
    return "\r\n".join(fold_line(line) for line in lines) + "\r\n"
