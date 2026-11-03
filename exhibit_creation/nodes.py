"""
Exhibit-creation nodes. No LLM: every step is deterministic.

  load_template -> fetch_data -> check_data -(unchanged?)-> build_exhibit -> persist
"""
from __future__ import annotations

import hashlib
import io
import json
from datetime import datetime, timezone
from pathlib import Path

from pptx import Presentation

import config
from core import alias_engine as eng
from core.alias_schema import AliasDocument
from core.exhibit_catalog import exhibit_folder, exhibit_name
from core.fidelity import compare
from core.metadata_extractor import extract_metadata
from core.response_schema import check_against_schema
from core.template_catalog import TemplateCatalog
from core.template_key import build_template_key
from exhibit_creation.state import ExhibitError
from integrations.datasources import build_datasource


def load_template(state, svc):
    try:
        key = build_template_key({"asset_type": state["asset_type"], "acronym": state["acronym"]})
    except ValueError as e:
        raise ExhibitError(str(e)) from e
    h = TemplateCatalog(svc.repos).find(state["asset_type"], state["acronym"], state.get("template_version"))
    if h is None:
        which = f"v{state['template_version']} of " if state.get("template_version") else ""
        raise ExhibitError(f"No template {which}{key.key}. Create it with the template workflow first.")
    svc.emit("node", f"Template {key.key} v{h.version} ({len(h.aliases.aliases)} aliases)")
    return {"key": key.as_dict(), "exhibit_name": exhibit_name(key.fields), "folder": exhibit_folder(key.as_dict()),
            "template": {"version": h.version, "location": h.location, "datasource_id": h.datasource_id},
            "aliases": h.aliases.model_dump(exclude_none=True), "log": [f"template {key.key} v{h.version}"]}


def fetch_data(state, svc):
    ds = svc.repos.datasources.get(state["template"]["datasource_id"])
    try:
        data = build_datasource(ds).fetch(state.get("params") or {})
    except Exception as e:
        raise ExhibitError(f"Datasource '{ds['id']}' failed: {e}") from e
    sha = hashlib.sha256(json.dumps(data, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()
    svc.emit("node", f"Datasource '{ds['id']}' returned data ({sha[:10]})")
    return {"data": data, "data_sha256": sha, "log": [f"fetched {ds['id']}"]}


def check_data(state, svc):
    """Fail before touching the slide if the data can't fill the template."""
    doc = AliasDocument.model_validate(state["aliases"])
    data, errors = state["data"], []
    for name, d in doc.aliases.items():
        try:
            vals = eng.resolve(d.path, data)
        except Exception as e:
            errors.append(f"{{{{{name}}}}}: invalid path {d.path!r} ({e})"); continue
        if not vals:
            errors.append(f"{{{{{name}}}}} ({d.description or d.path}): {d.path!r} finds nothing in the response")
        elif len(vals) > 1 and d.join is None:
            errors.append(f"{{{{{name}}}}}: {d.path!r} returns {len(vals)} values but the template expects one")
    for c in doc.charts:
        for p in [c.categories] + [s.values for s in c.series]:
            if not eng.resolve(p, data):
                errors.append(f"chart {c.shape_id}: {p!r} finds nothing in the response")

    schema_warnings = []
    loc = state["template"]["location"]
    if svc.repos.templates.exists(loc, "response_schema.json"):
        schema = json.loads(svc.repos.templates.read(loc, "response_schema.json"))
        schema_warnings = [f"API structure changed: {x}" for x in check_against_schema(data, schema)]

    out = {"schema_warnings": schema_warnings, "log": [f"data check: {len(errors)} errors"]}
    if errors and config.EXHIBIT_STRICT:
        out.update(status="failed", error="The response cannot fill the template:\n" + "\n".join(errors))
        return out
    out["fill_warnings"] = errors

    entry = svc.repos.exhibit_registry.get(state["key"]["key"])
    if entry and not state.get("force"):
        last = entry["versions"][-1]
        if last["data_sha256"] == state["data_sha256"] and last["template_version"] == state["template"]["version"]:
            svc.emit("node", f"Same template version and same data as v{last['version']}: nothing to do")
            out.update(status="unchanged", location=last["location"], version=last["version"])
    return out


def build_exhibit(state, svc):
    work = Path(state["work_dir"])
    loc = state["template"]["location"]
    doc = AliasDocument.model_validate(state["aliases"])
    prs = Presentation(io.BytesIO(svc.repos.templates.read(loc, "template.pptx")))
    warnings = list(state.get("fill_warnings", [])) + eng.fill(prs, doc, state["data"])
    left = eng.tokens_in(prs)
    if left:
        raise ExhibitError(f"Placeholders left unfilled: {sorted(left)}")

    name = state["exhibit_name"]
    prs.slides[0]._element.cSld.set("name", name)            # the slide's own name
    props = prs.core_properties
    props.title, props.subject = name, state["key"]["key"]
    props.keywords = f"{state['key']['key_hash']};template v{state['template']['version']}"
    out = work / f"{name}.pptx"
    prs.save(str(out))

    fidelity = {"status": "skipped", "fails": 0, "warnings": 0, "issues": []}
    if svc.repos.templates.exists(loc, "template_metadata.json"):
        ref = json.loads(svc.repos.templates.read(loc, "template_metadata.json"))
        markers = json.loads(svc.repos.templates.read(loc, "markers.json")) \
            if svc.repos.templates.exists(loc, "markers.json") else {}
        fidelity = compare(ref, extract_metadata(Presentation(str(out))),
                           locked_runs=markers.get("locked_runs", []), mode="exhibit")
    if fidelity["status"] == "fail":
        raise ExhibitError(f"Fidelity check failed ({fidelity['fails']} issues): the filled slide lost formatting")

    preview = None
    if svc.renderer:
        png = svc.renderer(out)
        if png:
            preview = str(work / "preview.png"); Path(preview).write_bytes(png)
    svc.emit("node", f"Built {name}.pptx; fidelity {fidelity['status']}; {len(warnings)} warning(s)")
    return {"exhibit_path": str(out), "preview_path": preview, "fidelity": fidelity, "fill_warnings": warnings,
            "log": [f"built {name}.pptx"]}


def persist(state, svc):
    key = state["key"]
    entry = svc.repos.exhibit_registry.get(key["key"])
    version = (entry["current_version"] + 1) if entry else 1
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    file_name = f"{state['exhibit_name']}.pptx"
    meta = {
        "exhibit_name": state["exhibit_name"], "key": key["key"], "key_hash": key["key_hash"],
        "asset_type": key["fields"].get("asset_type"), "acronym": key["fields"].get("acronym"),
        "version": version, "file_name": file_name, "generated_at": now,
        "template": state["template"], "params": state.get("params") or {},
        "data_sha256": state["data_sha256"], "fidelity": state["fidelity"]["status"],
        "warnings": state.get("fill_warnings", []) + state.get("schema_warnings", []),
    }
    files = {file_name: Path(state["exhibit_path"]).read_bytes(),
             "exhibit.json": json.dumps(meta, indent=2, ensure_ascii=False),
             "data.json": json.dumps(state["data"], indent=2, ensure_ascii=False),
             "fidelity_report.json": json.dumps(state["fidelity"], indent=2)}
    if state.get("preview_path"):
        files["preview.png"] = Path(state["preview_path"]).read_bytes()
    location = svc.repos.exhibits.save_version(state["folder"], version, files)
    rec = {"version": version, "location": location, "file_name": file_name, "generated_at": now,
           "template_version": state["template"]["version"], "data_sha256": state["data_sha256"],
           "params": state.get("params") or {}, "fidelity": state["fidelity"]["status"],
           "warnings": len(meta["warnings"])}
    svc.repos.exhibit_registry.upsert({
        "key": key["key"], "key_hash": key["key_hash"], "fields": key["fields"],
        "exhibit_name": state["exhibit_name"], "folder": state["folder"],
        "datasource_id": state["template"]["datasource_id"],
        "current_version": version, "current_location": location,
        "created_at": (entry or {}).get("created_at", now), "updated_at": now,
        "versions": (entry or {}).get("versions", []) + [rec]})
    status = "updated" if entry else "created"
    svc.emit("node", f"{state['exhibit_name']} {status}: v{version} at {location}")
    return {"location": location, "version": version, "status": status, "log": [f"{status} v{version}"]}
