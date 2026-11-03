"""Optional slide rendering via LibreOffice (used to give the LLM a picture of the Platinum copy)."""
from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

CANDIDATES = ["soffice", "libreoffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"]


def find_soffice() -> str | None:
    for c in CANDIDATES:
        found = shutil.which(c) or (c if Path(c).exists() else None)
        if found:
            return found
    return None


def render_png(pptx_path: Path, timeout: int = 120) -> bytes | None:
    """PNG of the first slide, or None when LibreOffice isn't installed (the workflow still runs)."""
    exe = find_soffice()
    if not exe:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        try:
            subprocess.run([exe, "--headless", "--convert-to", "png", "--outdir", tmp, str(pptx_path)],
                           check=True, capture_output=True, timeout=timeout)
        except Exception:
            return None
        pngs = list(Path(tmp).glob("*.png"))
        return pngs[0].read_bytes() if pngs else None
