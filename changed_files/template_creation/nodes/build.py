"""Nodes: build_template (placeholders + fidelity check + filled preview), persist (repository + registry)."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pptx import Presentation

from core import alias_engine as eng
from core.alias_schema import AliasDocument
from core.fidelity import compare
from core.metadata_extractor import extract_metadata as snapshot
from core.pptx_writer import remove_shapes
from template_creation import rules


def build_template(state, svc):
    work = Path(state["work_dir"])
    doc = AliasDocument.model_validate(state["alias_doc"])
    prs = Presentation(state["platinum_path"])
    removed_ids = rules.removed_shape_ids(doc, state["markers"])
    removed = remove_shapes(prs, removed_ids)
    n = eng.apply(prs, doc)
    out = work / "template.pptx"
    prs.save(str(out))

    reference = json.loads(Path(state["metadata_path"]).read_text(encoding="utf-8"))
    tmd = snapshot(Presentation(str(out)))
    (work / "template_metadata.json").write_text(json.dumps(tmd, indent=2), encoding="utf-8")
    fidelity = compare(reference, tmd, removed_ids=removed_ids,
                       locked_runs=state["markers"]["locked_runs"], mode="template")

    preview = Presentation(str(out))                      # what Agent 2 would produce with today's data
    fill_warnings = eng.fill(preview, doc, state["sample"])
    preview_path = work / "preview_filled.pptx"
    preview.save(str(preview_path))
    svc.emit("node", f"Template built with {n} placeholders; fidelity {fidelity['status']} "
                     f"({fidelity['fails']} fail, {fidelity['warnings']} warn)")
    result = {"template_path": str(out), "template_metadata_path": str(work / "template_metadata.json"),
              "preview_filled_path": str(preview_path), "fidelity": fidelity, "removed_shapes": removed,
              "placeholders": n, "validation_warnings": state.get("validation_warnings", []) + fill_warnings,
              "log": [f"built template ({n} placeholders), fidelity {fidelity['status']}"]}
    if fidelity["status"] == "fail":
        result.update(status="fidelity_failed",
                      error=f"Fidelity check failed with {fidelity['fails']} issue(s); template not saved")
    return result


def persist(state, svc):
    entry = state.get("registry_entry")
    version = (entry["current_version"] + 1) if entry else 1
    key = state["key"]
    doc = {**state["alias_doc"], "version": version}
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    files = {
        "template.pptx": Path(state["template_path"]).read_bytes(),
        "alias.json": json.dumps(doc, indent=2, ensure_ascii=False),
        "preview_filled.pptx": Path(state["preview_filled_path"]).read_bytes(),
        "platinum.pptx": Path(state["platinum_path"]).read_bytes(),
        "metadata.json": Path(state["metadata_path"]).read_text(encoding="utf-8"),
        "template_metadata.json": Path(state["template_metadata_path"]).read_text(encoding="utf-8"),
        "markers.json": json.dumps({**state["markers"], "decisions": doc.get("callouts", [])}, indent=2),
        "fidelity_report.json": json.dumps(state["fidelity"], indent=2),
        "response_schema.json": json.dumps(state["response_schema"], indent=2),
        "sample_response.json": json.dumps(state["sample"], indent=2, ensure_ascii=False),
        "story.json": json.dumps(state["story"], indent=2, ensure_ascii=False),
        "build_report.json": json.dumps({"placeholders": state["placeholders"], "removed": state["removed_shapes"],
                                         "diff": state.get("alias_diff"),
                                         "warnings": state.get("validation_warnings", [])}, indent=2),
    }
    location = svc.repos.templates.save_version(key["template_id"], version, files)
    version_rec = {"version": version, "location": location, "story_key": state["story"]["key"],
                   "platinum_sha256": state["platinum_sha256"], "created_at": now,
                   "fidelity": state["fidelity"]["status"], "placeholders": state["placeholders"]}
    new_entry = {
        "key": key["key"], "key_hash": key["key_hash"], "fields": key["fields"],
        "template_id": key["template_id"], "datasource_id": state["datasource_id"],
        "current_version": version, "current_location": location,
        "created_at": (entry or {}).get("created_at", now), "updated_at": now,
        "versions": (entry or {}).get("versions", []) + [version_rec],
    }
    svc.repos.registry.upsert(new_entry)
    status = "updated" if entry else "created"
    svc.emit("node", f"{key['key']} {status}: v{version} at {location}")
    return {"location": location, "version": version, "status": status, "log": [f"{status} v{version}"]}
