"""Exact, shared roster-column matching for preview and import.

Do not use substring matching: Company Name is not the attendee's Name and
Emergency Contact is not their Phone Number. Unknown columns belong to the event.
"""
import re


def normalize_header(value):
    return re.sub(r"[\s_]+", " ", str(value or "").strip().lower())


ALIASES = {
    "name": ("name", "full name", "fullname", "attendee name", "visitor name", "guest name"),
    "reg_no": ("registration number", "registration no", "reg number", "reg no", "regid", "ticket number", "ticket no", "id"),
    "email": ("email address", "email", "mail"),
    "phone": ("phone number", "phone", "mobile number", "mobile", "whatsapp", "contact number", "contact no", "contact"),
}


def match_columns(columns):
    """Return original column names, preferring canonical headers over aliases."""
    lookup = {normalize_header(c): c for c in columns}
    return {field: next(lookup[a] for a in aliases if a in lookup)
            for field, aliases in ALIASES.items() if any(a in lookup for a in aliases)}


def is_identity_column(column):
    return any(normalize_header(column) in aliases for aliases in ALIASES.values())


def validate_headers(columns):
    """Reject ambiguous input instead of silently dropping duplicate columns."""
    normalized = [normalize_header(c) for c in columns]
    if not normalized or any(not c for c in normalized):
        raise ValueError("Every roster column needs a header.")
    if len(set(normalized)) != len(normalized):
        raise ValueError("Duplicate column headers found. Give each column a unique name.")
