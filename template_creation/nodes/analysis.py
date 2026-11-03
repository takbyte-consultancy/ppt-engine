"""Nodes: extract_metadata (metadata + markers + preview), fetch_data."""
from __future__ import annotations

import json
from pathlib import Path

from pptx import Presentation

from core.markers import detect_markers
from core.metadata_extractor import extract_metadata as snapshot
from core.response_schema import infer_schema
from integrations.datasources import build_datasource
from template_creation.state import WorkflowError


def extract_metadata(state, svc):
    work = Path(state["work_dir"])
    prs = Presentation(state["platinum_path"])
    md = snapshot(prs)
    markers = detect_markers(prs)
    (work / "metadata.json").write_text(json.dumps(md, indent=2), encoding="utf-8")
    (work / "markers.json").write_text(json.dumps(markers, indent=2), encoding="utf-8")
    preview = None
    if svc.renderer:
        png = svc.renderer(Path(state["platinum_path"]))
        if png:
            preview = str(work / "preview.png")
            Path(preview).write_bytes(png)
    svc.emit("node", f"Metadata: {len(md['shapes'])} shapes; markers: {len(markers['callouts'])} callout(s), "
                     f"{len(markers['locked_runs'])} locked run(s)")
    return {"metadata_path": str(work / "metadata.json"), "markers": markers, "preview_path": preview,
            "original_names": {str(s["shape_id"]): s["name"] for s in md["shapes"]},
            "log": ["metadata + markers extracted"]}


def fetch_data(state, svc):
    ds = svc.repos.datasources.get(state["datasource_id"])
    try:
        data = build_datasource(ds).fetch(state.get("params") or {})
    except Exception as e:
        raise WorkflowError(f"Datasource '{ds['id']}' failed: {e}") from e
    svc.emit("node", f"Datasource '{ds['id']}' returned {len(json.dumps(data))} bytes")
    return {"sample": data, "response_schema": infer_schema(data), "log": [f"fetched {ds['id']}"]}
