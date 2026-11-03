"""
Fidelity check: proves the template (or a filled exhibit) kept the Platinum copy's formatting.

Compares two metadata snapshots shape-by-shape (by shape_id), ignoring text VALUES, shape names
(they were renamed on purpose) and shapes removed on purpose (instruction callouts + connectors).
Locked (red bold) runs are the exception: their text must match exactly too.

mode="template": paragraph counts and chart point counts must also match.
mode="exhibit" : those can legitimately change with the data, so they're not checked.
"""
from __future__ import annotations

TEXT_FRAME_KEYS = ("word_wrap", "auto_size", "vertical_anchor", "margins")
PARA_KEYS = ("level", "alignment", "space_before_pt", "space_after_pt", "line_spacing", "bullet",
             "margin_left", "indent")


class Report:
    def __init__(self):
        self.issues: list[dict] = []

    def add(self, severity: str, shape: str, path: str, expected, actual):
        self.issues.append({"severity": severity, "shape": shape, "path": path,
                            "expected": expected, "actual": actual})

    def result(self) -> dict:
        fails = sum(i["severity"] == "fail" for i in self.issues)
        warns = sum(i["severity"] == "warn" for i in self.issues)
        return {"status": "fail" if fails else "warn" if warns else "pass",
                "fails": fails, "warnings": warns, "issues": self.issues}


def _cmp(rep, shape, path, a, b, severity="fail"):
    if a != b:
        rep.add(severity, shape, path, a, b)


def _first_font(para):
    return para["runs"][0]["font"] if para["runs"] else None


def _cmp_paragraphs(rep, shape, path, pa, pb, mode):
    if len(pa) != len(pb):
        _cmp(rep, shape, f"{path}.paragraph_count", len(pa), len(pb),
             "fail" if mode == "template" else "warn")
    for i, (x, y) in enumerate(zip(pa, pb)):
        for k in PARA_KEYS:
            _cmp(rep, shape, f"{path}.p{i}.{k}", x.get(k), y.get(k))
        fx, fy = _first_font(x), _first_font(y)
        if fx and fy:
            _cmp(rep, shape, f"{path}.p{i}.font", fx, fy)
        distinct = {str(r["font"]) for r in x["runs"] if r["text"].strip()}
        if len(distinct) > 1 and len(y["runs"]) < len(x["runs"]):
            rep.add("warn", shape, f"{path}.p{i}.runs",
                    f"{len(x['runs'])} runs with mixed formatting", f"{len(y['runs'])} runs")


def _cmp_table(rep, shape, ta, tb, mode):
    for k in ("n_rows", "n_cols", "column_widths", "flags"):
        _cmp(rep, shape, f"table.{k}", ta[k], tb[k])
    _cmp(rep, shape, "table.row_heights", ta["row_heights"], tb["row_heights"],
         "fail" if mode == "template" else "warn")
    cb = {(c["r"], c["c"]): c for c in tb["cells"]}
    for ca in ta["cells"]:
        other = cb.get((ca["r"], ca["c"]))
        where = f"table.cell[{ca['r']},{ca['c']}]"
        if other is None:
            rep.add("fail", shape, where, "present", "missing"); continue
        for k in ("merge_origin", "spanned", "span", "fill", "margins", "vertical_anchor", "borders"):
            _cmp(rep, shape, f"{where}.{k}", ca[k], other[k])
        _cmp_paragraphs(rep, shape, where, ca["paragraphs"], other["paragraphs"], "exhibit")


def _cmp_chart(rep, shape, ca, cb, mode):
    for k in ("chart_type", "has_legend", "legend_position", "has_title", "value_axis"):
        _cmp(rep, shape, f"chart.{k}", ca[k], cb[k])
    _cmp(rep, shape, "chart.plot_count", len(ca["plots"]), len(cb["plots"]))
    for i, (pa, pb) in enumerate(zip(ca["plots"], cb["plots"])):
        for k in ("has_data_labels", "data_label_number_format"):
            _cmp(rep, shape, f"chart.plot{i}.{k}", pa[k], pb[k])
        _cmp(rep, shape, f"chart.plot{i}.series_count", len(pa["series"]), len(pb["series"]))
        if mode == "template":
            _cmp(rep, shape, f"chart.plot{i}.n_categories", pa["n_categories"], pb["n_categories"])
        for j, (sa, sb) in enumerate(zip(pa["series"], pb["series"])):
            for k in ("fill", "line"):
                _cmp(rep, shape, f"chart.plot{i}.series{j}.{k}", sa[k], sb[k])


def _locked_text(md_shape, lock):
    try:
        if "cell" in lock:
            r, c = lock["cell"]
            cell = next(x for x in md_shape["table"]["cells"] if x["r"] == r and x["c"] == c)
            return cell["paragraphs"][lock["paragraph"]]["runs"][lock["run"]]
        return md_shape["paragraphs"][lock["paragraph"]]["runs"][lock["run"]]
    except (KeyError, IndexError, StopIteration):
        return None


def compare(reference: dict, candidate: dict, *, removed_ids: set[int] = frozenset(),
            locked_runs: list[dict] | None = None, mode: str = "template") -> dict:
    rep = Report()
    ref = {s["shape_id"]: s for s in reference["shapes"] if s["shape_id"] not in removed_ids}
    cand = {s["shape_id"]: s for s in candidate["shapes"]}

    _cmp(rep, "<slide>", "slide.size", (reference["slide"]["width_emu"], reference["slide"]["height_emu"]),
         (candidate["slide"]["width_emu"], candidate["slide"]["height_emu"]))
    _cmp(rep, "<slide>", "slide.layout", reference["slide"]["layout"], candidate["slide"]["layout"])
    for sid in sorted(set(cand) - set(ref)):
        rep.add("fail", cand[sid]["name"], "shape", "absent", "unexpected extra shape")
    order_ref = [sid for sid in (s["shape_id"] for s in reference["shapes"]) if sid in ref and sid in cand]
    order_cand = [sid for sid in (s["shape_id"] for s in candidate["shapes"]) if sid in ref]
    _cmp(rep, "<slide>", "z_order", order_ref, order_cand)

    for sid, a in ref.items():
        b = cand.get(sid)
        label = f"{a['name']} (id {sid})"
        if b is None:
            rep.add("fail", label, "shape", "present", "missing"); continue
        label = f"{b['name']} (id {sid})"
        for k in ("shape_type", "auto_shape_type", "kind", "geometry_emu", "fill", "line"):
            _cmp(rep, label, k, a.get(k), b.get(k))
        if "text_frame" in a and "text_frame" in b:
            for k in TEXT_FRAME_KEYS:
                _cmp(rep, label, f"text_frame.{k}", a["text_frame"].get(k), b["text_frame"].get(k))
            _cmp_paragraphs(rep, label, "text", a["paragraphs"], b["paragraphs"], mode)
        if a.get("kind") == "table" and b.get("kind") == "table":
            _cmp_table(rep, label, a["table"], b["table"], mode)
        if a.get("kind") == "chart" and b.get("kind") == "chart":
            _cmp_chart(rep, label, a["chart"], b["chart"], mode)
        if a.get("kind") == "picture":
            _cmp(rep, label, "image.sha1", a["image"]["sha1"], (b.get("image") or {}).get("sha1"))

    for lock in locked_runs or []:
        sid = lock["shape_id"]
        if sid not in ref or sid not in cand:
            continue
        ra, rb = _locked_text(ref[sid], lock), _locked_text(cand[sid], lock)
        where = f"locked_run{lock.get('cell', '')}.p{lock['paragraph']}.r{lock['run']}"
        if rb is None:
            rep.add("fail", cand[sid]["name"], where, ra and ra["text"], "missing")
        else:
            _cmp(rep, cand[sid]["name"], where + ".text", ra["text"], rb["text"])
            _cmp(rep, cand[sid]["name"], where + ".font", ra["font"], rb["font"])
    return rep.result()
