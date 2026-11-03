# PPT-AI UI and Pitchbook Workflow Plan

## Purpose

This document captures the current project shape and the proposed UI and pitchbook addition so work can continue in another session without changing the existing agentic workflows by accident.

The requested product experience is a practical PPT-AI business workspace, informed by the supplied UI design reference. The main feature addition is a pitchbook creation step that follows exhibit creation. Existing template and exhibit logic stays intact; pitchbook assembly, review, feedback, and export are additional capabilities.

## What the project does today

The repository is a Python/Streamlit application for generating data-driven PowerPoint templates and exhibits:

1. **Template creation (Agent 1)** takes a Platinum/mock-up PowerPoint and story metadata, extracts formatting and component metadata, retrieves data, maps dynamic text to data paths, validates the mapping, pauses for human approval, builds a reusable template, and stores a version.
2. **Exhibit creation (Agent 2)** is a deterministic LangGraph workflow. It loads a template by asset type and acronym, fetches datasource output, checks that values fit the template, fills the slide, checks visual/content fidelity, and saves a versioned exhibit. The exhibit workflow does not use an LLM.
3. **Current UI** is `app.py`, a Streamlit UI for template creation, template registry, and datasource registration/inspection. It is not yet the PPT-AI dashboard described in the design reference, and it does not expose the exhibit workflow or pitchbook creation.

The README identifies a later Book assembly agent as the next step. This plan develops that next step into a user-facing pitchbook workflow.

## Existing code and storage map

| Area | Current responsibility | Pitchbook relevance |
|---|---|---|
| `README.md` | Product purpose, architecture, setup, and existing workflow documentation | Source of truth for current behavior |
| `app.py` | Streamlit template workflow, review UI, registry, and datasource tabs | Current UI entry point; pitchbook UI can be added alongside existing workflow UI without moving agent logic into the interface |
| `template_creation/` | LangGraph workflow, state, nodes, services, runner, mapping and approval | Existing template generation; leave behavior unchanged |
| `exhibit_creation/graph.py` | LangGraph wiring for the deterministic exhibit workflow | Existing exhibit graph; do not alter its sequence or logic for pitchbook support |
| `exhibit_creation/nodes.py` | Template lookup, datasource fetch, validation, slide fill, fidelity check, persistence | Produces the single-slide artifacts that pitchbooks will assemble |
| `exhibit_creation/runner.py` | `ExhibitCreationWorkflow.run(...)` facade | Existing public boundary for creating an exhibit |
| `core/exhibit_catalog.py` | `ExhibitCatalog.find/search`, `ExhibitHandle`, slide bytes/metadata/preview access | Existing read API for locating saved slides by asset type/acronym and latest or explicit version |
| `core/template_key.py` | Normalizes asset type/acronym keys and computes the key hash | Stable identity building block; the existing hash is key identity, while exhibit versions have their own version numbers and locations |
| `repositories/base.py` | Repository protocols and `Repositories` container | Exhibit registry and file repository are already exposed; pitchbook storage can be added behind repository interfaces |
| `repositories/json_store/__init__.py` | JSON registries and immutable versioned file storage | Current backend for exhibit artifacts; future pitchbook manifests/outputs can follow this versioned pattern |
| `tests/` | Existing workflow tests and test fakes | Add pitchbook-focused coverage when implementation begins; preserve current tests and behavior |

### Existing exhibit identity and lookup

Exhibit creation derives a normalized key from configured key fields (normally asset type + acronym). A key hash is included in the exhibit folder name. Each saved exhibit registry entry records its key, key hash, fields, current location, and version history. Each version has its own version number, storage location, filename, template version, data hash, and metadata. `ExhibitCatalog` can locate the latest or a requested version by asset type and acronym and can search the latest exhibit per key.

There is no separate first-class `UID` field in the current code. For the pitchbook feature, define an explicit stable exhibit reference that resolves through the catalog/registry, for example `{key, version}` (and optionally `key_hash` as a verified identity field). Do not assume the key hash by itself identifies a particular slide version. Sample pitchbook references should resolve to the intended saved exhibit version, with unresolved or ambiguous references reported for review.

## Product UI direction

Build a responsive PPT-AI workspace rather than a marketing landing page. Keep the original agentic workflows and their orchestration intact; the UI should call workflow facades and catalogs, not LangGraph nodes or storage internals directly.

The workspace should provide:

- Branded header with PPT-AI identity, concise product label, API docs link, and light/dark toggle.
- Session context for tenant, user, role, and authentication status.
- A clear workflow navigation/task flow for template creation, exhibit creation, and the new pitchbook creation step.
- Document upload/indexing and document Q&A with cited sources, plus natural-language workflow request and status feedback, as described in the supplied design prompt where supported by product capabilities.
- Responsive agent cards with concise descriptions, relevant inputs, and clear run actions.
- Consistent status pills for ready, running, needs approval, completed, and failed; show progress and errors beside the relevant task.
- Deep navy and blue-violet accents, restrained gradients, readable typography, subtle borders, accessible contrast, and calm spacing. Keep dark and light themes legible, including the logo.

The document search, upload, and generic agent-catalog elements in the design prompt are not represented by current repository code. Treat them as UI/product expansion requiring separate data and service integrations; do not imply that the existing template/exhibit workflows already provide those capabilities.

## Pitchbook user flow

Add a **Pitchbook creation** step after exhibit creation. The pitchbook workflow consumes saved exhibit slides and produces a multi-slide PowerPoint draft, supports human review and feedback, then produces a downloadable final presentation.

### 1. Choose how to assemble

Offer two clear modes:

1. **Use a pitchbook sample**: user supplies/selects a sample deck or manifest that contains intended exhibit/acronym references and slide positions. Parse each reference into a stable exhibit key and version/UID, resolve it against `ExhibitCatalog`, and place the matching exhibit slide into the indicated pitchbook position. Show missing, duplicate, or ambiguous references before continuing.
2. **Create from a request**: user enters a plain-English brief, such as an acronym, financial year, and desired deck scope. A LangGraph pitchbook workflow interprets the criteria, queries the exhibit catalog for candidate slides, and creates an ordered draft proposal. Criteria should be visible/editable; the system must not silently invent unavailable exhibits.

### 2. Human review and ordering

Before final assembly, present the selected slide list with slide previews, exhibit identity, version, title/metadata, and proposed order. Let the reviewer reorder slides (drag/drop or explicit move-to-position actions), remove/add eligible exhibits, and confirm the draft. Persist order as explicit slide entries in the pitchbook manifest.

### 3. Draft deck and feedback

Assemble the selected exhibit slides into a draft PPTX while preserving each slide's formatting. Send the draft to a feedback/revision step. Collect feedback at three scopes:

- **Pitchbook-level**: deck naming, overall requirements, and slide order.
- **Slide-level**: a comment attached to a particular pitchbook slide.
- **Component-level**: a comment attached to a particular shape/component on a slide.

Store comments as structured feedback records linked to a pitchbook draft ID, slide identity, and (for component feedback) a stable component locator such as source exhibit key/version plus shape ID and a fallback locator. PowerPoint comments/annotations can be emitted in the exported deck where technically reliable, but the feedback record remains the system of record so comments do not get lost during PPTX transformations. Reviewers can ask the feedback agent to apply requested changes; the agent should route edits to the relevant deck, slide, or component target and report what it could not apply.

The interface should allow an explicit **Resolve all** action after feedback is addressed. It must show the pending/resolved state per comment and avoid marking feedback resolved until the user confirms the result.

### 4. Finalize and download

After feedback is resolved and the user approves the final draft, save an immutable pitchbook version with its manifest, PPTX, and feedback/revision history. Provide a direct PPTX download and show the final status and version. Keep draft and final states distinct.

## Proposed architecture (additive)

Introduce a separate `pitchbook_creation/` package rather than modifying the existing template/exhibit graph logic. A first implementation can include:

- `pitchbook_creation/state.py`: JSON-serializable LangGraph state for source mode, criteria, candidates, slide order, feedback, output, and status.
- `pitchbook_creation/graph.py`: orchestration for resolve/select -> assemble draft -> human review/feedback -> revise -> finalize/persist. Approval pauses should use the project's established runner/checkpoint conventions where feasible.
- `pitchbook_creation/nodes.py`: catalog resolution, manifest/sample parsing, draft assembly, feedback application, and persistence. Keep deterministic actions deterministic; use an LLM only for interpreting natural-language selection/feedback when necessary, with human approval before consequential changes.
- `pitchbook_creation/runner.py`: facade for UI/API callers, following the existing `TemplateCreationWorkflow` and `ExhibitCreationWorkflow` boundary pattern.
- `core/pitchbook_catalog.py` (or equivalent): read-side lookup for saved pitchbook drafts/final versions.
- Repository interfaces for pitchbook registry/artifact storage, implemented first with the JSON/file backend to match the current architecture.
- UI step/tab in `app.py` that calls the pitchbook runner and catalog, with selection, preview, ordering, feedback, approval, and download states.

Use `python-pptx` capabilities where possible, but verify the source slide copying approach carefully: the existing code reads/writes PowerPoint presentations, while robustly copying slides between presentations can require lower-level OOXML handling or a dedicated helper. Isolate this in a tested assembly utility. Do not mutate persisted exhibit source files; treat them as immutable inputs.

## Suggested implementation sequence

1. **Confirm data contract**: define sample/manifest format, stable slide reference (`key` + version), pitchbook statuses, feedback scopes, and what a component locator looks like.
2. **Catalog/selection layer**: expose all needed exhibit versions and metadata, including deterministic resolution and validation errors.
3. **Assembly utility**: combine selected exhibit slides in requested order while preserving shapes, formatting, notes where supported, and source provenance.
4. **Pitchbook workflow**: add isolated LangGraph state/nodes/runner for source resolution, draft creation, human review, feedback, and final persistence.
5. **UI**: add the pitchbook step to the product workspace and connect it only through runner/catalog APIs; apply the PPT-AI visual direction consistently.
6. **Validation**: cover reference resolution, ordering, missing slides, PPTX assembly, feedback targeting, resolve-all behavior, versioned persistence, and regression of existing template/exhibit workflows.

## Guardrails

- Preserve the existing template and exhibit graph behavior and public runner contracts.
- The pitchbook workflow is an additional downstream workflow, not a replacement or rewrite of existing agents.
- Never infer that a key hash alone points to a specific exhibit version; persist an explicit versioned reference.
- Keep exhibit artifacts immutable; assemble copies into a new pitchbook artifact.
- Human approval controls final slide selection/order and feedback resolution.
- Keep a provenance map from each pitchbook slide/component back to its source exhibit and version.
- Do not advertise document Q&A, document indexing, or unrelated agent capabilities until their backend integrations exist.

## Files inspected for this plan

- `README.md`
- `app.py`
- `exhibit_creation/graph.py`
- `exhibit_creation/state.py`
- `exhibit_creation/nodes.py`
- `exhibit_creation/runner.py`
- `core/exhibit_catalog.py`
- `core/template_key.py`
- `repositories/base.py`
- `repositories/json_store/__init__.py`
- `config.py`

## Implementation progress

The additive first implementation is now in the repository:

- `app.py` has PPT-AI workspace branding, a light/dark selector, and workspace steps for template creation, exhibit creation, and Pitchbook Studio alongside registry/datasource screens.
- `pitchbook_creation/` contains a separate LangGraph resolver, draft assembler, human review interrupt, constrained feedback agent, and versioned persistence flow.
- The supplied `LiveDoc_New Business Book_GB_FIT_EU_EN_Z_USD.PPT.PPTX` is a 49-slide, A4-landscape reference deck. Its content slides contain visible identifiers such as `GB.…`, `SAS.…`, `FI.…`, and `ITC.…`; it also includes section dividers, dense tables, images, grouped shapes, and embedded charts. Sample mode extracts these visible identifiers and records each mapping to a versioned exhibit UID (`key_hash:vN`). The sample's existing slide content is reference data and is not copied into the new output.
- Request mode accepts acronym, asset type, and financial-year criteria; without manually selected references, the LangGraph resolver selects matching latest versions.
- Review supports explicit slide positions, a renamed final deck, and pitchbook/slide/component feedback. The feedback agent proposes constrained deck rename, slide move, and component replacement actions when `OPENAI_API_KEY` is configured; exact component text replacements also work without a key. A second human approval screen shows proposed revisions before final persistence. Feedback is recorded in the manifest and unresolved notes remain open.
- `core/pptx_assembler.py` copies editable slide shapes, image relationships, and recursively related OOXML parts such as charts/workbooks into a new deck. The sample has 30 masters and 37 layouts, so cross-master styling and notes still need review against representative outputs.
- `config.py` now defines `EXHIBITS_DIR` (the repository factory already referenced it) and `PITCHBOOKS_DIR`; storage interfaces include pitchbook registry/artifact repositories. `README.md` describes the additional workflow.

No existing template or exhibit graph nodes were changed. The app is running at `http://127.0.0.1:8501`. The workflow code has not been tested in this implementation pass. A next session should validate assembly against representative exhibit PPTX files, especially charts, grouped shapes, themes, and slide notes, then add pitchbook-specific regression coverage.
