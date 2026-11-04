# Next Plan: Automated Pitchbook Creation and Collaborative Review

**Status:** Draft for discussion  
**Scope:** Product and architecture proposal only. No implementation changes are included in this document.

## 1. Goal

Let a user describe a financial pitchbook once, then use registered templates, datasource integrations, and saved exhibit versions to prepare a reviewable PowerPoint. Use an optional LLM to interpret requests and feedback, while deterministic code resolves data, fills slides, assembles the deck, and validates changes. Keep a human approval step before finalization.

The system must not invent financial values, silently select ambiguous source material, or change slide elements outside the requested scope.

## 2. Current Project State

The project already has three separate workflow areas:

1. **Template creation:** analyzes an uploaded Platinum slide or story, maps dynamic content to datasource paths, validates the mapping, and saves a versioned reusable template. The mapping stage can use an LLM.
2. **Exhibit creation:** deterministically loads a template, retrieves datasource values, fills the slide, validates it, and saves a versioned exhibit.
3. **Pitchbook creation:** can resolve saved exhibits from a sample deck or a request, assemble slides, collect pitchbook/slide/component feedback, and save an approved deck.

Relevant project components include:

- `core/template_catalog.py` and `core/exhibit_catalog.py` for template and exhibit lookup.
- `exhibit_creation/` for deterministic datasource retrieval, filling, validation, and persistence.
- `pitchbook_creation/` for request resolution, assembly, review, feedback, and persistence.
- `repositories/` for the current JSON registry and versioned file storage.
- `app.py` for the current Streamlit interface.

Current limitations to account for:

- Request mode has basic acronym, asset type, and financial-year filters. It does not yet turn a full brief into a complete, explainable plan.
- No first-class request schema currently represents the full user brief, assumptions, missing information, and proposed ordering.
- The app currently uses local JSON/file storage and has no configured user authentication or shared-review permissions.
- Pitchbook graph state uses an in-memory checkpointer, so paused work is lost when the process restarts.
- PPTX merging is sensitive to slide dimensions, layouts, masters, themes, backgrounds, charts, and other OOXML relationships. Visual fidelity needs ongoing regression checks against source exhibit previews.

## 3. Terms and Identity

Keep these concepts distinct in code, UI, and manifests:

- **Template:** reusable slide design with placeholders and datasource aliases.
- **Exhibit:** a generated, data-filled slide from a specific template version and datasource response.
- **Pitchbook:** an ordered collection of specific exhibit versions.
- **Exhibit UID:** a stable exhibit identity plus an immutable version. The current convention is based on the exhibit key hash and version, such as `key_hash:vN`.

The current normalized key is generally based on **asset type + acronym**. Acronym alone is not a safe identity. Folder names are storage details, not the product-level reference. Every pitchbook slide should record the exhibit key, key hash, exact version, template version, and relevant datasource snapshot/provenance.

## 4. Proposed End-to-End User Flow

### Step 1: One natural-language request

The user provides a brief such as:

> Create the FY2026 EUR pitchbook for IMNASMA. Include the latest income and risk exhibits, put the overview first, and exclude the legacy page.

The UI should accept optional constraints for asset type, acronym, region, currency, fiscal period, eligible source deck/sample, ordering hints, and output title. The user should not need to understand internal storage keys.

### Step 2: Build a structured, reviewable plan

An optional LLM converts the brief into a typed plan containing:

- normalized asset type and acronym;
- fiscal period, currency, region, and other requested filters;
- required and excluded exhibit concepts;
- ordering constraints and suggested order;
- desired deck title;
- unresolved or ambiguous requirements;
- a confidence/explanation record tied to actual catalog evidence.

The model can only choose from catalog results and supported operations. It must not invent a UID, template, exhibit, datasource field, or value. If required information is missing or multiple candidates fit, ask the user to choose before execution.

### Step 3: Resolve existing exhibits and identify gaps

Resolve each requested exhibit through the catalogs using the normalized identity and available metadata. Pin all selected exhibits to exact versions before assembly.

For each requested item, classify it as:

- available as a saved exhibit;
- eligible to generate from an existing template and datasource;
- blocked because a template, datasource, parameter, or value is missing;
- ambiguous and awaiting user selection.

Show these results before running generation. Never silently substitute a different acronym, year, currency, region, or version.

### Step 4: Generate only missing exhibits

When the required template and datasource are registered, call the existing Exhibit workflow with the requested parameters. The exhibit agent remains deterministic: retrieve data, fill mapped values, run validation/fidelity checks, and save an immutable version.

If data cannot be retrieved, a required alias cannot be resolved, or validation fails, stop that exhibit and present the specific reason. Do not fill gaps using model-generated financial values.

### Step 5: Assemble a draft deterministically

Once all required exhibits are resolved, assemble their saved slide artifacts in the proposed order. Use a sample deck or an approved manifest as a source of ordering/layout constraints where provided. Preserve source exhibit provenance and do not modify the saved exhibit files.

The model may recommend order from explicit user instructions and approved examples. It does not author slide content or directly rewrite PPTX XML.

### Step 6: Human review of the draft

Show a large, readable preview of the actual assembled draft, with:

- slide thumbnails and a full-size viewer;
- proposed order and reorder controls;
- exhibit name, UID, exact version, and source template;
- visible missing/omitted items and relevant warnings;
- component selection with a stable source locator;
- approve, revise, and reject actions.

Changing order, removing a slide, or selecting another exhibit must update the draft preview before approval.

### Step 7: Feedback and constrained revision

Reviewers can comment at three levels:

- **Pitchbook:** title, overall scope, ordering, or global instruction.
- **Slide:** request that applies to one selected slide.
- **Component:** request attached to a particular shape, text run, table cell, or chart element where supported.

The LLM can translate a vague comment into an operation proposal, for example:

- `move_slide(source_uid, position)`;
- `rename_pitchbook(title)`;
- `replace_component_text(source_uid, component_locator, exact_text)`;
- `regenerate_exhibit(source_uid, datasource_parameters)` when the requested change is data-driven and the template supports it.

Every operation must pass application-side validation against the current pitchbook, selected exhibit version, component locator, and available datasource schema. Show the proposed operation, target, before/after values, and preview to the reviewer. Apply only approved operations. Unsupported or ambiguous feedback stays open with a clear explanation.

For financial values or reporting dates, prefer updating the datasource request and regenerating the affected exhibit from its template. This preserves the design and avoids editing unrelated PPTX elements. A direct text replacement should only be used when the reviewer explicitly requests literal text and the target is unambiguous.

### Step 8: Final approval and immutable version

After revision, show a final diff and preview. A reviewer explicitly approves the final deck and may resolve comments. Save a new immutable version containing:

- the PPTX;
- a pitchbook manifest with ordered exhibit UIDs and provenance;
- request criteria and resolved assumptions;
- feedback threads, actions, approvals, and resolution state;
- validation and visual-fidelity results;
- model/provider metadata for AI-assisted decisions, when enabled.

The user can then download the final deck.

## 5. LLM Boundary and Guardrails

### Appropriate LLM roles

- Parse user briefs and free-text comments into structured intents.
- Identify ambiguity and ask focused questions.
- Suggest ordering using user constraints and approved historical examples.
- Match natural-language exhibit concepts to candidates returned by the catalog.
- Explain why an item is unavailable, only using supplied system evidence.

### Deterministic responsibilities

- Catalog and UID lookup.
- Datasource calls and data provenance.
- Exhibit filling and fidelity validation.
- PPTX assembly, order application, exact text replacement, and persistence.
- Authorization checks, action validation, approval state, and audit history.

### Safety rules

- Give the model a finite catalog of candidate exhibit UIDs and allowed operations.
- Use strict structured output and validate every field in application code.
- Require exact targets for component edits; reject edits if the locator is stale or ambiguous.
- Reject fabricated or unavailable identifiers, years, currencies, values, and fields.
- Never give the LLM a generic “rewrite this deck” action.
- Preserve untouched source components by applying validated changes to a copy of the immutable exhibit source.
- Require human approval for the assembled draft and for revisions before finalization.
- Record unresolved requests; do not mark comments resolved just because an agent attempted a change.

OpenAI function calling supports schema-described tools, while the application remains responsible for executing and validating tool actions. See [Function calling](https://developers.openai.com/api/docs/guides/function-calling) and [Structured outputs](https://developers.openai.com/api/docs/guides/structured-outputs).

## 6. Historical Examples and Model Training

Do not start with fine-tuning or a local model. Start with:

1. Existing template alias documents, registries, and datasource schemas as authoritative structured sources.
2. Approved pitchbook manifests and ordering patterns as examples for retrieval, filtered by asset type, acronym family, region, and period.
3. An explicit record of which examples influenced a recommendation, so the reviewer can inspect it.
4. A labeled evaluation set of requests, expected source matches, valid operation plans, and known ambiguous/unsupported cases.

Use previous PPTX files as **retrieval examples and layout references**, not as unverified truth. Historical decks may be stale, contain different scopes, or have been manually edited. Their data should never override the current datasource or selected exhibit version.

Consider fine-tuning only after collecting enough reviewed intent-to-operation examples and demonstrating a repeatable quality gap that prompt/schema improvements cannot address. Consider a local model only if privacy, deployment, and latency requirements justify its maintenance cost and it meets the same evaluation bar.

OpenAI recommends using representative datasets and explicit evaluation criteria to check model behavior; see [Evaluation best practices](https://developers.openai.com/api/docs/guides/evaluation-best-practices).

## 7. Collaborative Review Dashboard

The shared dashboard should provide:

- authenticated reviewer access with roles and deck-level permissions;
- a versioned deck viewer with slide navigation and full-size render;
- comment pins or a side panel mapped to slide/component locators;
- threaded replies, open/resolved states, assignee, and audit history;
- a proposed-change panel showing before/after values and the affected source exhibit;
- re-preview after every approved change;
- explicit final approval and downloadable artifact.

Keep the dashboard comment system as the source of truth. PowerPoint itself supports slide comments via PresentationML/Open XML, but interoperability and component-level anchoring across PowerPoint variants must be validated before relying on embedded comments as the canonical record. See Microsoft’s [Open XML presentation overview](https://learn.microsoft.com/en-us/office/open-xml/presentation/overview) and [modern slide comment example](https://learn.microsoft.com/en-us/office/open-xml/presentation/how-to-add-a-comment-to-a-slide-in-a-presentation).

The current Streamlit app has no configured authentication or shared review service. A production multi-reviewer dashboard therefore needs an identity provider, API/backend, database-backed comment and pitchbook state, object storage for immutable PPTX versions, and access-controlled share links.

## 8. Persistence and Deployment

Replace ephemeral pitchbook/checkpoint state with durable storage before enabling real cross-user approval flows. The project currently uses `InMemorySaver`; LangGraph documents in-memory checkpointing as suitable for debugging/testing and describes persistent savers such as SQLite and Postgres for longer-lived workflows. See [LangGraph checkpointing](https://langchain-ai.github.io/langgraph/reference/checkpoints/) and [interrupts/HITL](https://langchain-ai.github.io/langgraph/how-tos/human_in_the_loop/breakpoints/).

Recommended production data model:

- relational database for users, requests, pitchbook runs, slide lists, feedback threads, operations, approvals, and audit events;
- durable LangGraph checkpointer keyed by tenant/user/workflow thread;
- object storage for source PPTX artifacts, previews, and immutable final versions;
- registries retaining stable exhibit keys and version references;
- idempotency keys for datasource calls, generation attempts, and revision application.

## 9. Suggested Delivery Phases

### Phase A: Request contract and resolver

- Define request, candidate, ordering, assumption, and unresolved-reason schemas.
- Resolve against current template and exhibit registries with exact version pins.
- Add a reviewable plan screen before generation.

### Phase B: End-to-end deterministic generation

- Wire request criteria into existing datasource and exhibit workflow interfaces.
- Generate only missing exhibits with existing registered templates.
- Create a manifest and assemble a draft deterministically.
- Add stop conditions for absent templates, missing datasource values, and validation failures.

### Phase C: Preview and feedback operations

- Improve full-deck preview and per-slide component selection.
- Convert comments into a validated, finite operation plan.
- Add before/after diffs and approval for each revision batch.

### Phase D: Shared review and durable state

- Add identity, sharing permissions, persistent workflow checkpoints, comment threads, and audit records.
- Add version retention and recovery.
- Validate export and review behavior with representative Microsoft PowerPoint files.

### Phase E: Evaluation and optimization

- Evaluate request parsing, exhibit resolution, order proposals, feedback action selection, and unsupported-request behavior.
- Measure catalog match accuracy, unsupported action rate, wrong-target rate, visual-fidelity regression, and human correction rate.
- Only then consider fine-tuning or local inference.

## 10. Acceptance Criteria

- A single brief can produce an explainable plan that lists its criteria, source exhibits, versions, and unresolved items.
- Every selected slide resolves to an existing exhibit UID/version or a newly validated exhibit generated through the existing workflow.
- Missing or ambiguous data blocks generation with a specific explanation; the system never fabricates financial content.
- All slide ordering and feedback operations are visible, schema-valid, target-validated, and reversible before final approval.
- Unrequested slide components remain unchanged; source exhibit versions remain immutable.
- Visual regression checks compare each assembled slide against its source exhibit preview at the target canvas size.
- Reviewers can comment and approve from a shared, permission-controlled dashboard, and the workflow survives app/process restarts.
- Final exports retain a manifest, source provenance, feedback history, and approval state.

## 11. Questions for Product Discussion

1. Which fields are required in the initial natural-language request: acronym, asset type, period, currency, region, or client?
2. Should the system generate any missing exhibit automatically when a valid template and datasource exist, or show a proposed generation list for confirmation first?
3. Which prior decks are approved examples for ordering, and who marks them as trusted/current?
4. Which component operations are in scope for the first release: exact text replacement, datasource-backed regeneration, table-cell edits, chart data refresh, or only comments/order/name?
5. Who can create share links, invite reviewers, resolve comments, and give final approval?
6. Must feedback be embedded in the exported PPTX, or is the authenticated review dashboard the canonical comment record?
7. Which data classes may be sent to a hosted LLM, and should AI feedback be disabled per deck or tenant?

