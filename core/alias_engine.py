"""
Alias engine: everything about placeholders, deterministic and shared by both agents.

  propose()   scan the Platinum copy against the datasource and decide, per text fragment:
                1. red bold (locked)                 -> never touched
                2. exactly equal to a JSON value      -> alias the whole text
                3. equal to a JSON key name           -> static label
                4. part of the text equals a value    -> alias that part, keep the rest literal
                5. otherwise                          -> static
              Ambiguous values (same value at several paths) are resolved from table row/column labels
              or the words just before the value; what stays ambiguous is flagged low-confidence.
  validate()  every alias resolves, every occurrence still matches the slide, nothing data-backed is
              silently left static, and a round-trip (apply -> fill with the same data) reproduces the slide
  apply()     write {{alias}} placeholders into a presentation (template creation)
  fill()      replace placeholders with formatted data (exhibit creation, Agent 2)
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Any

from jsonpath_ng.ext import parse as jp_parse
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.shapes import MSO_SHAPE_TYPE

from core.alias_schema import (AliasDef, AliasDocument, KeptStatic, Location, Occurrence, StaticLabel)
from core.pptx_inspector import iter_shapes
from core.pptx_writer import replace_span

TOKEN = re.compile(r"\{\{\s*([A-Za-z0-9_.\-]+)\s*\}\}")
EDGE = "\u2014\u2013-\u2022\u00b7 :;,.()%"
VALUE_LIKE = re.compile(r"^[\s+\-\u2014\u2013]*[\d.,]*\s*%?\s*$")
NUM_FORMATS = ["{value}", "{value:.2f}", "{value:.1f}", "{value:.0f}", "{value:,.2f}", "{value:,.0f}",
               "{value:.3f}", "{value:+.2f}", "{value:.2f}%", "{value:.1f}%", "{value}%"]
IDENT = re.compile(r"[A-Za-z_][A-Za-z0-9_]*$")


def normws(s) -> str:
    return re.sub(r"\s+", " ", str(s).replace("\x0b", " ")).strip()


def norm_label(s) -> str:
    s = normws(s).lower().replace("_", " ").strip(EDGE)
    return re.sub(r"\s+", " ", s)


# =========================================================================== #
# JSON side
# =========================================================================== #
@dataclass
class Leaf:
    path: str
    value: Any
    segments: list[str]          # normalised key names along the path


def _seg(k: str) -> str:
    return f".{k}" if IDENT.match(k) else "['" + k.replace("'", "\\'") + "']"


def flatten(data) -> tuple[list[Leaf], dict[str, str]]:
    """All leaves (strings, numbers, nulls) with standard JSONPaths, and {normalised key name: path}."""
    leaves: list[Leaf] = []
    keys: dict[str, str] = {}

    def rec(v, path, segs):
        if isinstance(v, dict):
            for k, x in v.items():
                p = path + _seg(str(k))
                keys.setdefault(norm_label(k), p)
                rec(x, p, segs + [norm_label(k)])
        elif isinstance(v, list):
            for i, x in enumerate(v):
                if isinstance(x, dict) and isinstance(x.get("shape_name"), str):
                    p = f"{path}[?(@.shape_name=='" + x["shape_name"].replace("'", "\\'") + "')]"
                else:
                    p = f"{path}[{i}]"
                rec(x, p, segs)
        else:
            leaves.append(Leaf(path, v, segs))

    rec(data, "$", [])
    return leaves, keys


def resolve(path: str, data) -> list:
    return [m.value for m in jp_parse(path).find(data)]


def parent_path(path: str) -> str:
    """Path of the object holding this leaf: drop the last .key / ['key'] / [i] / [?(...)] segment."""
    if path.endswith("']"):
        return path[:path.rfind("['")]
    if path.endswith("]"):
        depth, i = 0, len(path) - 1
        while i >= 0:
            depth += {"]": 1, "[": -1}.get(path[i], 0)
            if depth == 0:
                return path[:i]
            i -= 1
    return path[:path.rfind(".")] if "." in path else path


def describe(path: str) -> str:
    return " > ".join(p for p in re.split(r"\.(?![^\[]*\])", path.lstrip("$.")) if p)


# =========================================================================== #
# Slide side
# =========================================================================== #
@dataclass
class Unit:
    shape_id: int
    shape_name: str
    loc: Location
    text: str
    run_spans: list[tuple[int, int]]
    locked: list[tuple[int, int]]
    table: int | None = None
    row: int | None = None
    col: int | None = None

    @property
    def key(self):
        return (self.shape_id, tuple(self.loc.cell) if self.loc.cell else None, self.loc.paragraph)


def para_text(p) -> str:
    return "".join(r.text for r in p.runs)


def _run_spans(p) -> list[tuple[int, int]]:
    out, pos = [], 0
    for r in p.runs:
        out.append((pos, pos + len(r.text))); pos += len(r.text)
    return out


def get_paragraph(prs, shape_id: int, loc: Location):
    shp = next(s for s in iter_shapes(prs.slides[0].shapes) if s.shape_id == shape_id)
    tf = shp.table.cell(*loc.cell).text_frame if loc.cell else shp.text_frame
    return tf.paragraphs[loc.paragraph]


def table_text_grid(table) -> list[list[str]]:
    """Cell texts where spanned cells carry their merge-origin's text (for row/column context)."""
    n_r, n_c = len(table.rows), len(table.columns)
    grid = [[table.cell(r, c).text for c in range(n_c)] for r in range(n_r)]
    for r in range(n_r):
        for c in range(n_c):
            cell = table.cell(r, c)
            if cell.is_merge_origin:
                for rr in range(r, r + cell.span_height):
                    for cc in range(c, c + cell.span_width):
                        grid[rr][cc] = cell.text
    return grid


def units(prs, markers: dict, exclude_ids: set[int]) -> tuple[list[Unit], dict]:
    locked_runs: dict[tuple, set[int]] = {}
    for lk in markers.get("locked_runs", []):
        k = (lk["shape_id"], tuple(lk["cell"]) if "cell" in lk else None, lk["paragraph"])
        locked_runs.setdefault(k, set()).add(lk["run"])
    fully = set(markers.get("fully_locked_shape_ids", []))
    locked_cells = {(int(k), tuple(c)) for k, v in markers.get("locked_cells", {}).items() for c in v}

    out, grids = [], {}
    for s in iter_shapes(prs.slides[0].shapes):
        if s.shape_id in exclude_ids or s.shape_id in fully or s.shape_type in (MSO_SHAPE_TYPE.GROUP,
                                                                                 MSO_SHAPE_TYPE.PICTURE):
            continue
        if getattr(s, "has_chart", False):
            continue
        frames = []
        if getattr(s, "has_table", False) and s.has_table:
            grids[s.shape_id] = table_text_grid(s.table)
            for r, row in enumerate(s.table.rows):
                for c, cell in enumerate(row.cells):
                    if not cell.is_spanned and (s.shape_id, (r, c)) not in locked_cells:
                        frames.append((cell.text_frame, [r, c]))
        elif getattr(s, "has_text_frame", False) and s.has_text_frame:
            frames.append((s.text_frame, None))
        for tf, cell in frames:
            for pi, p in enumerate(tf.paragraphs):
                text = para_text(p)
                if not text.strip():
                    continue
                spans = _run_spans(p)
                lr = locked_runs.get((s.shape_id, tuple(cell) if cell else None, pi), set())
                out.append(Unit(s.shape_id, s.name, Location(paragraph=pi, cell=cell), text, spans,
                                [spans[i] for i in lr if i < len(spans)],
                                s.shape_id if cell else None, cell[0] if cell else None, cell[1] if cell else None))
    return out, grids


# =========================================================================== #
# Matching
# =========================================================================== #
@dataclass
class Match:
    start: int
    end: int
    candidates: list[tuple[Leaf, str | None]]      # (leaf, number format)
    kind: str                                      # exact | case_insensitive | substring | structural


def _boundary(text, s, e) -> bool:
    before = text[s - 1] if s > 0 else " "
    after = text[e] if e < len(text) else " "
    return not (before.isalnum() or before == "_") and not (after.isalnum() or after == "_")


def _code_like(v: str) -> bool:
    return " " not in v and len(v) >= 4 and bool(re.search(r"[_\d]", v))


def number_format_for(v, text: str) -> str | None:
    """Format reproducing the slide text; prefers the slide's decimal places so 0.4 prints as 0.40."""
    core = text.rstrip("%")
    dec = len(core.split(".")[-1]) if "." in core else 0
    pct = "%" if text.endswith("%") else ""
    preferred = [f"{{value:.{dec}f}}{pct}", f"{{value:,.{dec}f}}{pct}"] if dec else [f"{{value}}{pct}", f"{{value:,.0f}}{pct}"]
    for f in preferred + NUM_FORMATS:
        try:
            if f.format(value=v) == text:
                return f
        except (ValueError, TypeError):
            pass
    return None


def case_rule(slide: str, value: str) -> str | None:
    slide, value = normws(slide), normws(value)
    if slide == value:
        return None
    for rule, fn in (("lower", str.lower), ("upper", str.upper),
                     ("lower_first", lambda v: v[:1].lower() + v[1:]),
                     ("upper_first", lambda v: v[:1].upper() + v[1:])):
        if fn(value) == slide:
            return rule
    return None


def _overlaps(a, b) -> bool:
    return a[0] < b[1] and b[0] < a[1]


class Matcher:
    def __init__(self, data):
        self.leaves, self.keys = flatten(data)
        self.strings: dict[str, list[Leaf]] = {}
        self.numbers: list[Leaf] = []
        for lf in self.leaves:
            if isinstance(lf.value, str) and lf.value.strip():
                self.strings.setdefault(normws(lf.value).lower(), []).append(lf)
            elif isinstance(lf.value, (int, float)) and not isinstance(lf.value, bool):
                self.numbers.append(lf)

    def match_unit(self, u: Unit) -> tuple[list[Match], str | None]:
        """Returns (matches, label_key_path). Implements the precedence rules."""
        t = u.text
        stripped = t.strip()
        s0 = t.index(stripped)
        whole = (s0, s0 + len(stripped))
        if any(_overlaps(whole, lk) for lk in u.locked) and len(u.locked) == len(u.run_spans):
            return [], None
        key = normws(stripped).lower()

        exact = [lf for lf in self.strings.get(key, []) if normws(lf.value) == normws(stripped)]
        nums = [(lf, f) for lf in self.numbers if (f := number_format_for(lf.value, stripped))]
        if (exact or nums) and not u.locked:                                       # rule 2
            return [Match(*whole, [(lf, None) for lf in exact] + nums, "exact")], None
        if norm_label(stripped) in self.keys:                                      # rule 3
            return [], self.keys[norm_label(stripped)]
        if key in self.strings and not u.locked:                                   # rule 2b (case)
            return [Match(*whole, [(lf, None) for lf in self.strings[key]], "case_insensitive")], None

        found: list[Match] = []                                                    # rule 4
        tl = t.lower()
        for v, lfs in sorted(self.strings.items(), key=lambda kv: -len(kv[0])):
            if not (len(v) >= 8 or _code_like(lfs[0].value)) or norm_label(v) in self.keys:
                continue
            pos = tl.find(v)
            while pos >= 0:
                span = (pos, pos + len(v))
                if (_boundary(t, *span) and not any(_overlaps(span, (m.start, m.end)) for m in found)
                        and not any(_overlaps(span, lk) for lk in u.locked)):
                    found.append(Match(*span, [(lf, None) for lf in lfs], "substring"))
                pos = tl.find(v, pos + 1)
        return sorted(found, key=lambda m: m.start), None


def _context(u: Unit, grids: dict, start: int) -> tuple[str, set[str]]:
    if u.table is not None:
        g = grids[u.table]
        labels = {norm_label(g[u.row][c]) for c in range(len(g[0])) if c != u.col}
        labels |= {norm_label(g[r][u.col]) for r in range(len(g)) if r != u.row}
        return "table", {x for x in labels if x}
    before = re.split(r"[.!?]\s", u.text[:start])[-1]
    return "text", set(re.findall(r"[a-z0-9]+", before.lower())[-10:])


def _score(lf: Leaf, kind: str, ctx: set[str]) -> int:
    if kind == "table":
        return sum(1 for seg in lf.segments if seg in ctx)
    toks = re.findall(r"[a-z0-9]+", lf.segments[-1]) if lf.segments else []
    return sum(1 for t in toks if t in ctx)


def _choose(m: Match, u: Unit, grids: dict, affinity: dict | None = None):
    kind, ctx = _context(u, grids, m.start)
    scored = []
    for lf, fmt in m.candidates:
        sc = _score(lf, kind, ctx)
        if affinity:
            sc += affinity.get(parent_path(lf.path), 0) * 0.1
        scored.append((sc, -len(lf.segments), lf, fmt))
    scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
    best = scored[0]
    if len(scored) == 1:
        conf = "high"
    elif scored[1][0] < best[0]:
        conf = "medium"
    else:
        conf = "low"                       # tied on evidence; path length only broke the tie
    return best[2], best[3], conf, [s[2].path for s in scored[1:]]


# =========================================================================== #
# Proposal
# =========================================================================== #
def _alias_names():
    import itertools
    import string
    for n in itertools.count(1):
        for combo in itertools.product(string.ascii_uppercase, repeat=n):
            yield "".join(combo)


def _def_key(d: AliasDef) -> tuple:
    return (d.path, d.format, d.case, d.null, d.join)


def name_aliases(doc: AliasDocument, previous: AliasDocument | None = None) -> None:
    """Give every distinct definition a short alias, reusing names from the previous version."""
    prev = {_def_key(d): name for name, d in (previous.aliases.items() if previous else [])}
    defs: dict[tuple, AliasDef] = {}
    for o in doc.occurrences:
        d = doc.aliases.get(o.alias)
        if d is not None:
            defs.setdefault(_def_key(d), d)
    old_alias_of = {o.id: _def_key(doc.aliases[o.alias]) for o in doc.occurrences if o.alias in doc.aliases}
    used = {prev[k] for k in defs if k in prev}
    gen = (n for n in _alias_names() if n not in used and n not in (previous.aliases if previous else {}))
    new_names = {}
    for o in doc.occurrences:                               # reading order -> A, B, C...
        k = old_alias_of.get(o.id)
        if k is not None and k not in new_names:
            new_names[k] = prev.get(k) or next(gen)
    doc.aliases = {new_names[k]: defs[k] for k in defs if k in new_names}
    for o in doc.occurrences:
        if o.id in old_alias_of:
            o.alias = new_names[old_alias_of[o.id]]


@dataclass
class Proposal:
    doc: AliasDocument
    unresolved: list[dict]


def propose(prs, data, markers: dict, exclude_ids: set[int] | None = None,
            previous: AliasDocument | None = None) -> Proposal:
    exclude_ids = exclude_ids or set()
    m = Matcher(data)
    us, grids = units(prs, markers, exclude_ids)
    doc = AliasDocument()
    pending: list[tuple[Unit, Match]] = []
    covered: dict[tuple, list[tuple[int, int]]] = {}
    for u in us:
        matches, label = m.match_unit(u)
        if label:
            doc.labels.append(StaticLabel(shape_id=u.shape_id, location=u.loc, text=u.text.strip(), key_path=label))
        for mt in matches:
            pending.append((u, mt))
            covered.setdefault(u.key, []).append((mt.start, mt.end))

    def choose_all():
        ch = {i: _choose(mt, u, grids) for i, (u, mt) in enumerate(pending)}
        rows: dict[tuple, dict] = {}
        for i, (u, mt) in enumerate(pending):
            if u.table is not None:
                par = parent_path(ch[i][0].path)
                d = rows.setdefault((u.table, u.row), {}); d[par] = d.get(par, 0) + 1
        for i, (u, mt) in enumerate(pending):
            if ch[i][2] == "low" and u.table is not None:
                ch[i] = _choose(mt, u, grids, rows.get((u.table, u.row)))
        return ch, rows

    chosen, row_parents = choose_all()

    # structural pass: value-looking table cells with no match (nulls, changed numbers)
    table_formats: dict[int, list[str]] = {}
    for i, (u, mt) in enumerate(pending):
        if chosen[i][1] and u.table is not None:
            table_formats.setdefault(u.table, []).append(chosen[i][1])
    for u in us:
        if u.table is None or u.key in covered or not VALUE_LIKE.match(u.text) \
                or not re.search(r"[\d\u2014\u2013-]", u.text):
            continue
        _, ctx = _context(u, grids, 0)
        aff = row_parents.get((u.table, u.row), {})
        scored = sorted(((_score(lf, "table", ctx) + 0.1 * aff.get(parent_path(lf.path), 0), lf)
                         for lf in m.leaves if not isinstance(lf.value, str)), key=lambda x: x[0], reverse=True)
        if scored and scored[0][0] >= 2 and (len(scored) == 1 or scored[1][0] < scored[0][0]):
            lf = scored[0][1]
            stripped = u.text.strip()
            if re.search(r"\d", stripped):
                dec = len(stripped.split(".")[1]) if "." in stripped else 0
                fmt = f"{{value:.{dec}f}}"
            elif table_formats.get(u.table):
                fmt = max(set(table_formats[u.table]), key=table_formats[u.table].count)
            else:
                fmt = None
            s0 = u.text.index(stripped)
            mt = Match(s0, s0 + len(stripped), [(lf, fmt)], "structural")
            mt.null_display = stripped if lf.value is None else ""
            pending.append((u, mt))
            chosen[len(pending) - 1] = (lf, fmt, "medium", [])
            covered.setdefault(u.key, []).append((mt.start, mt.end))

    for i, (u, mt) in enumerate(pending):
        lf, fmt, conf, others = chosen[i]
        frag = u.text[mt.start:mt.end]
        d = AliasDef(path=lf.path, format=fmt,
                     case=case_rule(frag, lf.value) if isinstance(lf.value, str) else None,
                     null=getattr(mt, "null_display", ""), description=describe(lf.path), sample=lf.value)
        how = {"exact": "identical to the data value", "case_insensitive": "same as the data value (case differs)",
               "substring": "data value inside the text; the rest stays literal",
               "structural": "row and column labels match this data field"}[mt.kind]
        tmp = f"_{i}"
        doc.aliases[tmp] = d
        doc.occurrences.append(Occurrence(id=f"o{i + 1}", alias=tmp, shape_id=u.shape_id, shape_name=u.shape_name,
                                          location=u.loc, start=mt.start, end=mt.end, text=frag, how=how,
                                          confidence=conf if mt.kind != "structural" else "medium",
                                          candidates=others[:4]))
    name_aliases(doc, previous)

    unresolved = []
    for u in us:
        spans = covered.get(u.key, [])
        rest = "".join(ch if not any(s <= i < e for s, e in spans) else " " for i, ch in enumerate(u.text))
        if re.search(r"\d", rest) and u.key not in {(lb.shape_id, tuple(lb.location.cell) if lb.location.cell else None,
                                                     lb.location.paragraph) for lb in doc.labels}:
            unresolved.append({"shape_id": u.shape_id, "shape_name": u.shape_name, "location": u.loc.model_dump(),
                               "text": u.text, "not_covered": normws(rest)})
    return Proposal(doc, unresolved)


# =========================================================================== #
# Editing helpers (used by the agent's tools and by reviewers)
# =========================================================================== #
def _assign(doc: AliasDocument, occ: Occurrence, d: AliasDef, previous: AliasDocument | None) -> None:
    tmp = f"_new{len(doc.aliases)}"
    doc.aliases[tmp] = d
    occ.alias = tmp
    name_aliases(doc, previous)


def set_definition(doc, occ_id: str, d: AliasDef, previous=None) -> Occurrence:
    occ = next(o for o in doc.occurrences if o.id == occ_id)
    _assign(doc, occ, d, previous)
    return occ


def add_occurrence(doc, prs, shape_id: int, loc: Location, text: str, d: AliasDef, nth: int = 1,
                   how: str = "added by the mapping agent", previous=None) -> Occurrence:
    para = para_text(get_paragraph(prs, shape_id, loc))
    pos, start = -1, -1
    for _ in range(nth):
        pos = para.find(text, pos + 1)
        if pos < 0:
            raise ValueError(f"'{text}' (occurrence {nth}) not found in that paragraph: {para[:120]!r}")
        start = pos
    span = (start, start + len(text))
    for o in doc.occurrences:
        if o.shape_id == shape_id and o.location == loc and _overlaps(span, (o.start, o.end)):
            raise ValueError(f"overlaps existing placeholder {o.id} ({o.text!r}); change that one instead")
    shp = next(s for s in iter_shapes(prs.slides[0].shapes) if s.shape_id == shape_id)
    occ = Occurrence(id=f"o{max([int(o.id[1:]) for o in doc.occurrences if o.id[1:].isdigit()] + [0]) + 1}",
                     alias="", shape_id=shape_id, shape_name=shp.name, location=loc, start=span[0], end=span[1],
                     text=text, how=how, confidence="medium")
    doc.occurrences.append(occ)
    _assign(doc, occ, d, previous)
    return occ


def keep_static(doc, occ_id: str, reason: str, previous=None) -> None:
    if not reason.strip():
        raise ValueError("a reason is required to keep data-backed text static")
    occ = next(o for o in doc.occurrences if o.id == occ_id)
    doc.occurrences.remove(occ)
    doc.kept_static.append(KeptStatic(shape_id=occ.shape_id, location=occ.location, text=occ.text, reason=reason))
    name_aliases(doc, previous)


# =========================================================================== #
# Values
# =========================================================================== #
def _apply_case(s: str, rule: str | None) -> str:
    if rule == "lower":
        return s.lower()
    if rule == "upper":
        return s.upper()
    if rule == "lower_first":
        return s[:1].lower() + s[1:]
    if rule == "upper_first":
        return s[:1].upper() + s[1:]
    return s


def _one(v, d: AliasDef) -> str:
    if v is None:
        return d.null
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return d.format.format(value=v) if d.format else str(v)
        except (ValueError, TypeError):
            return str(v)
    return _apply_case(str(v), d.case)


def value_for(d: AliasDef, data) -> str:
    vals = resolve(d.path, data)
    if d.join is not None:
        return d.join.join(_one(v, d) for v in vals)
    return _one(vals[0] if vals else None, d)


# =========================================================================== #
# Apply (template) and fill (exhibit)
# =========================================================================== #
def apply(prs, doc: AliasDocument) -> int:
    """Replace each occurrence's Platinum text with {{alias}} (right-to-left per paragraph)."""
    groups: dict[tuple, list[Occurrence]] = {}
    for o in doc.occurrences:
        groups.setdefault((o.shape_id, tuple(o.location.cell or ()), o.location.paragraph), []).append(o)
    n = 0
    for occs in groups.values():
        p = get_paragraph(prs, occs[0].shape_id, occs[0].location)
        for o in sorted(occs, key=lambda o: o.start, reverse=True):
            replace_span(p, o.start, o.end, "{{" + o.alias + "}}"); n += 1
    return n


def _all_paragraphs(prs):
    for s in iter_shapes(prs.slides[0].shapes):
        if getattr(s, "has_table", False) and s.has_table:
            for row in s.table.rows:
                for cell in row.cells:
                    yield from cell.text_frame.paragraphs
        elif getattr(s, "has_text_frame", False) and s.has_text_frame:
            yield from s.text_frame.paragraphs


def tokens_in(prs) -> dict[str, int]:
    out: dict[str, int] = {}
    for p in _all_paragraphs(prs):
        for m in TOKEN.finditer(para_text(p)):
            out[m.group(1)] = out.get(m.group(1), 0) + 1
    return out


def fill(prs, doc: AliasDocument, data) -> list[str]:
    """Agent 2: replace every {{alias}} with its formatted value. Tokens split across runs are handled."""
    warnings = []
    for p in _all_paragraphs(prs):
        text = para_text(p)
        for m in reversed(list(TOKEN.finditer(text))):
            d = doc.aliases.get(m.group(1))
            if d is None:
                warnings.append(f"unknown placeholder {m.group(0)}"); continue
            try:
                value = value_for(d, data)
            except Exception as e:
                warnings.append(f"{m.group(0)}: {e}"); value = ""
            replace_span(p, m.start(), m.end(), value)
    shapes = {s.shape_id: s for s in iter_shapes(prs.slides[0].shapes)}
    for ch in doc.charts:
        cd = CategoryChartData()
        cd.categories = resolve(ch.categories, data)
        for s in ch.series:
            cd.add_series(s.name, resolve(s.values, data))
        shapes[ch.shape_id].chart.replace_data(cd)
    return warnings


def _copy(prs):
    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return Presentation(buf)


# =========================================================================== #
# Validation
# =========================================================================== #
def validate(doc: AliasDocument, prs, data, markers: dict, exclude_ids: set[int],
             removed_ids: set[int]) -> tuple[list[str], list[str]]:
    errors, warnings = [], []

    used = {o.alias for o in doc.occurrences}
    for name, d in doc.aliases.items():
        if name not in used:
            errors.append(f"alias {name} is not used by any placeholder"); continue
        try:
            vals = resolve(d.path, data)
        except Exception as e:
            errors.append(f"alias {name}: invalid JSONPath {d.path!r} ({e})"); continue
        if not vals:
            errors.append(f"alias {name}: {d.path!r} finds nothing in the data (JSONPath is case-sensitive)")
        elif len(vals) > 1 and d.join is None:
            errors.append(f"alias {name}: {d.path!r} returns {len(vals)} values; index it or set 'join'")

    us, _ = units(prs, markers, exclude_ids)
    unit_text = {u.key: u for u in us}
    spans: dict[tuple, list] = {}
    for o in doc.occurrences:
        if o.alias not in doc.aliases:
            errors.append(f"{o.id}: alias {o.alias} is not defined")
        if o.shape_id in removed_ids:
            errors.append(f"{o.id}: shape {o.shape_id} is an instruction callout and is removed from the template")
            continue
        k = (o.shape_id, tuple(o.location.cell) if o.location.cell else None, o.location.paragraph)
        u = unit_text.get(k)
        if u is None:
            errors.append(f"{o.id}: no editable text at shape {o.shape_id} {o.location.model_dump()} "
                          f"(missing, empty or red bold locked)"); continue
        if u.text[o.start:o.end] != o.text:
            errors.append(f"{o.id}: expected {o.text!r} at {o.start}-{o.end}, slide has {u.text[o.start:o.end]!r}")
        if any(_overlaps((o.start, o.end), lk) for lk in u.locked):
            errors.append(f"{o.id}: overlaps red bold (locked) text in {u.shape_name}")
        for other in spans.get(k, []):
            if _overlaps((o.start, o.end), (other.start, other.end)):
                errors.append(f"{o.id} overlaps {other.id} in {u.shape_name}")
        spans.setdefault(k, []).append(o)
        if o.confidence == "low":
            warnings.append(f"{o.id} {{{{{o.alias}}}}} {o.text[:40]!r}: ambiguous, chose {doc.aliases[o.alias].path}"
                            f" over {o.candidates[:2]}" if o.alias in doc.aliases else "")

    # coverage: data-backed text must be aliased or explicitly kept static
    fresh = propose(prs, data, markers, exclude_ids)
    kept = {(k.shape_id, tuple(k.location.cell) if k.location.cell else None, k.location.paragraph, normws(k.text))
            for k in doc.kept_static}
    for o in fresh.doc.occurrences:
        if o.confidence == "low":
            continue
        k = (o.shape_id, tuple(o.location.cell) if o.location.cell else None, o.location.paragraph)
        hit = any(_overlaps((o.start, o.end), (x.start, x.end)) for x in spans.get(k, []))
        if not hit and (*k, normws(o.text)) not in kept:
            errors.append(f"{o.text[:50]!r} in {o.shape_name} is in the data ({fresh.doc.aliases[o.alias].path}) "
                          f"but has no placeholder; alias it or keep_static with a reason")
    for k in doc.kept_static:
        warnings.append(f"kept static although data-backed: {k.text[:50]!r} (reason: {k.reason})")

    # charts
    bound = {c.shape_id for c in doc.charts}
    kept_ids = {k.shape_id for k in doc.kept_static}
    for s in iter_shapes(prs.slides[0].shapes):
        if getattr(s, "has_chart", False) and s.shape_id not in exclude_ids | removed_ids:
            if s.shape_id not in bound and s.shape_id not in kept_ids:
                errors.append(f"chart {s.name} (id {s.shape_id}) has no data binding (bind_chart or keep_static)")
    for c in doc.charts:
        for p in [c.categories] + [x.values for x in c.series]:
            try:
                if not resolve(p, data):
                    errors.append(f"chart {c.shape_id}: {p!r} finds nothing")
            except Exception as e:
                errors.append(f"chart {c.shape_id}: invalid JSONPath {p!r} ({e})")

    # round trip: template + same data must reproduce the Platinum copy's text
    if not errors:
        rt = _copy(prs)
        apply(rt, doc)
        fill(rt, doc, data)
        for u in us:
            try:
                filled = para_text(get_paragraph(rt, u.shape_id, u.loc))
            except Exception:
                continue
            if normws(filled) != normws(u.text):
                warnings.append(f"round trip changes text in {u.shape_name}: {normws(u.text)[:50]!r} -> "
                                f"{normws(filled)[:50]!r}")
    return errors, [w for w in warnings if w]
