# PPT-AI · PowerPoint workflow

Turns a Platinum copy (mock-up slide from a JIRA story) into a reusable, data-mapped PowerPoint
template, registered under a unique key (asset type + acronym) for the Exhibit agent to use.

## Workflow (LangGraph)

```
load_story -> extract_platinum -> resolve_key --(Platinum copy unchanged)--> END
                                       | create (new key) / update (key exists -> v n+1)
                                       v
extract_metadata -> fetch_data -> map_components <-> validate -> human_review -> build_template -> persist
 (metadata.json,     (JSON/Excel/    (LLM agent)    (retry with   (pause for     (strip data,      (template
  colour markers)     HTTP)                          feedback)     approval)      fidelity check)   repo + registry)
```

Only `map_components` uses an LLM. Detection of facts (colours, locks, geometry, formatting) is always code.

### Templates are the Platinum copy with placeholders

Wherever the datasource provides the text, the template holds a short placeholder (`{{A}}`, `{{B}}`, ...);
everything else stays exactly as written. `alias.json` maps every placeholder to a JSONPath (plus number
format, case and null display). Code proposes the placeholders, the LLM resolves what code cannot, a human
approves. Precedence per text fragment:

1. red bold -> locked, never changed
2. exactly equal to a JSON value -> placeholder
3. equal to a JSON key name ("Cumulative total returns") -> static label
4. part of the text equals a value -> placeholder for that part ("Data for {{S}} (%)")
5. anything else -> static

Validation: every path must resolve (case-sensitive; `['3_months']` for keys starting with a digit), no
data-backed text may silently stay static (`keep_static` needs a reason), and a round trip (template filled
with the same data) must reproduce the Platinum copy.

| Platinum copy convention | What happens |
|---|---|
| Yellow / green filled shape with text | Callout candidate. The LLM decides *instruction* (applied to its target, then the callout and its arrows are deleted) or *design element* (kept and mapped). |
| Red **bold** text | Locked. Kept exactly as-is (text and red bold) in the template and every exhibit. Whole shape locked -> `static`; mixed -> run-level `runs` mapping; table cell -> `locked_cells` (set by code). |

## Layout

```
integrations/
  stories.py              Story model + StorySource interface + LocalStorySource (uploads)
  jira/                   everything JIRA: client, story parser, JiraStorySource
  datasources/            DataSource interface + json_file / excel_file / http connectors
repositories/
  base.py                 storage interfaces + get_repositories()  <- the one switch for DB later
  json_store/             JSON-file implementations (data/*.json, templates/)
core/                     deterministic building blocks
  pptx_inspector.py       read slides, rename shapes, trim to one slide
  pptx_writer.py          format-preserving span replacement, shape removal
  alias_schema.py         alias.json contract (aliases, occurrences, labels, callouts, charts)
  alias_engine.py         propose / validate / apply placeholders / fill (Agent 2 uses fill)
  metadata_extractor.py   full formatting/locator snapshot -> metadata.json
  markers.py              theme-aware colour classification, callouts, connectors, red bold locks
  fidelity.py             compares two snapshots -> fidelity_report.json
  template_key.py         key = normalised asset type + acronym (configurable)
  template_catalog.py     read API for Agent 2: find(asset_type, acronym)
  llm_agent.py            OpenAI Responses-API tool loop
template_creation/
  graph.py                LangGraph wiring          runner.py   start / resume / result facade
  nodes/                  one file per step         mapper.py   Mapper interface + LLM implementation
  rules.py                callout validation, enrichment, version diff
  prompts/mapping_agent.md
app.py                    Streamlit UI              scripts/create_template.py   CLI
data/                     datasources.json, template_registry.json
templates/<template_id>/v<n>/   template.pptx, alias.json, preview_filled.pptx, metadata.json, template_metadata.json,
                                markers.json, fidelity_report.json, response_schema.json,
                                sample_response.json, story.json, build_report.json, platinum.pptx
```

## Setup

```bash
python -m venv .venv && .venv\Scripts\activate        # Mac/Linux: source .venv/bin/activate
pip install -r requirements.txt
copy .env.example .env                                 # Mac/Linux: cp; then set OPENAI_API_KEY / OPENAI_MODEL
python scripts/make_annotated_platinum.py              # test Platinum copy with every colour marker
```

Optional: install LibreOffice so the LLM also gets an image of the slide (better callout targeting).

## Run

```bash
python -m tests.test_template_workflow                 # full offline test suite, no keys needed
streamlit run app.py                                   # UI
python -m scripts.create_template --pptx mock/platinum_american_bond.pptx \
       --asset-type "Fixed Income" --acronym IMNASMA --datasource ds_american_bond   # CLI
```

## Configuration

* **JIRA**: set `JIRA_BASE_URL` + credentials. Asset type / acronym / datasource are read from the custom
  fields named in `JIRA_FIELD_*`, else from `Asset Type: …` / `Acronym: …` / `Datasource: …` lines in the
  description; the UI can override them. The newest `.pptx` attachment (preferring names containing
  "platinum") is used.
* **Template key**: `TEMPLATE_KEY_FIELDS=asset_type,acronym`. Values are normalised (`fixed-income` ==
  `Fixed Income`), so spelling differences can't create duplicate templates.
* **Datasources**: `data/datasources.json` or the UI's Datasources tab. Excel files become
  `{"sheets": {"<name>": {"grid": [[…]], "records": [{…}]}}}`.

## Extending

* **Database**: implement the three interfaces in `repositories/base.py` (e.g. `repositories/sql_store/`)
  and return them from `get_repositories()`. Nothing else changes.
* **Another story system** (Azure DevOps, ServiceNow…): implement `StorySource`.
* **Another model or agent framework**: implement `Mapper.map()`.
* **Exhibit agent**: `h = TemplateCatalog().find(asset_type, acronym)`; `prs = h.presentation()`;
  `core.alias_engine.fill(prs, h.aliases, data)`; verify with `core.fidelity.compare(..., mode="exhibit")`.

## Pitchbook creation (Agent 3)

The Streamlit workspace includes a Pitchbook Studio after exhibit creation. It accepts either a reference
pitchbook whose slides the user maps to saved exhibits, or a request with acronym/asset type/financial year
criteria. When explicit exhibits are not selected for a request, the LangGraph resolver selects matching
latest exhibit versions from the registry. The reviewer can reorder slides, add deck/slide/component feedback,
review feedback-agent revisions in a second human approval step, and download the saved PPTX. Pitchbook registry
entries retain a manifest with source exhibit key hashes and exact versions, order, criteria, and feedback.

Implementation lives in `pitchbook_creation/`; it is an additive downstream workflow and does not change the
template or exhibit graphs. Sample-deck mode extracts visible identifiers (for example `GB.…` and `SAS.…`)
and records each mapping to a versioned exhibit UID (`key_hash:vN`). `core/pptx_assembler.py` copies editable
shapes and recursively related OOXML parts such as charts and workbooks into a new presentation.
Component-level feedback supports exact text replacement; when `OPENAI_API_KEY` is configured, the feedback
agent can also interpret notes into constrained title, slide-order, and component-text actions. Unsupported
notes remain open with an explanation. Existing exhibit artifacts remain immutable.

Run with `streamlit run app.py`. The app creates `exhibits/` and `pitchbooks/` artifact directories and JSON
registries as needed. The supplied 49-slide sample is a visual and identifier reference; sample mode maps its
visible slide identifiers and does not copy the sample's existing slide content into the new pitchbook.
