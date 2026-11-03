"""
The template key: a unique, deterministic identifier computed from the configured key fields
(asset type + acronym by default). Used by the registry, by Agent 2 to find templates, and later
as the database primary key (key_hash).
"""
from __future__ import annotations

import hashlib
import re
from dataclasses import asdict, dataclass

import config


def normalise(value: str) -> str:
    v = re.sub(r"[\s\-/]+", "_", str(value).strip().upper())
    return re.sub(r"[^A-Z0-9_]", "", v).strip("_")


@dataclass(frozen=True)
class TemplateKey:
    key: str                  # FIXED_INCOME::IMNASMA          (human readable, registry key)
    key_hash: str             # 16 hex chars of sha256(key)    (stable id for the database)
    template_id: str          # fixed_income__imnasma          (folder name)
    fields: dict

    def as_dict(self) -> dict:
        return asdict(self)


def build_template_key(values: dict, fields: list[str] | None = None) -> TemplateKey:
    fields = fields or config.TEMPLATE_KEY_FIELDS
    normalised = {f: normalise(values.get(f, "")) for f in fields}
    missing = [f for f, v in normalised.items() if not v]
    if missing:
        raise ValueError(f"Cannot build template key, missing: {missing}")
    key = "::".join(normalised[f] for f in fields)
    return TemplateKey(key=key, key_hash=hashlib.sha256(key.encode()).hexdigest()[:16],
                       template_id="__".join(normalised[f].lower() for f in fields),
                       fields={f: str(values[f]).strip() for f in fields})
