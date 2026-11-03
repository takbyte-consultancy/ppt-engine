"""
Command-line runner (no UI). Pauses at review and asks in the terminal.

  python -m scripts.create_template --pptx mock/platinum_american_bond.pptx \
         --asset-type "Fixed Income" --acronym IMNASMA --datasource ds_american_bond
  python -m scripts.create_template --jira PB-123
"""
import argparse
import json
from pathlib import Path

from integrations.stories import LocalStorySource
from template_creation.runner import TemplateCreationWorkflow
from template_creation.services import default_services


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--jira"); ap.add_argument("--pptx")
    ap.add_argument("--asset-type", default=""); ap.add_argument("--acronym", default="")
    ap.add_argument("--datasource", default=""); ap.add_argument("--slide", type=int)
    ap.add_argument("--force", action="store_true"); ap.add_argument("--no-vision", action="store_true")
    a = ap.parse_args()
    if a.jira:
        from integrations.jira import JiraStorySource
        source, ref = JiraStorySource(), a.jira
    elif a.pptx:
        source, ref = LocalStorySource(Path(a.pptx), a.asset_type, a.acronym, a.datasource), f"LOCAL-{Path(a.pptx).stem}"
    else:
        ap.error("--jira or --pptx is required")

    log = lambda k, p: print(f"  {p}" if k == "node" else f"    · {p['tool']}" if k == "tool_call" else "", end="\n" if k in ("node", "tool_call") else "")
    wf = TemplateCreationWorkflow(default_services(source, log, use_vision=not a.no_vision))
    r = wf.start(ref, overrides={"asset_type": a.asset_type, "acronym": a.acronym, "datasource_id": a.datasource},
                 slide_index=(a.slide - 1) if a.slide else None, force=a.force)
    while r.status == "awaiting_review":
        rv = r.review
        print(f"\nREVIEW {rv['template_key']} ({rv['mode']}): {len(rv['mapping']['components'])} components, "
              f"{len(rv['callouts'])} callouts, {len(rv['locked_runs'])} locked runs")
        for w in rv["warnings"] + rv.get("errors", []):
            print("  !", w)
        Path("work/review_mapping.json").write_text(json.dumps(rv["mapping"], indent=2), encoding="utf-8")
        choice = input("Mapping written to work/review_mapping.json. [a]pprove / [e]dited file / [r]eject: ").strip().lower()
        if choice == "r":
            r = wf.resume(r.run_id, {"action": "reject", "reason": input("Reason: ")})
        elif choice == "e":
            r = wf.resume(r.run_id, {"action": "approve",
                                     "mapping": json.loads(Path("work/review_mapping.json").read_text(encoding="utf-8"))})
        else:
            r = wf.resume(r.run_id, {"action": "approve"})
    print(f"\nResult: {r.status}", r.error or r.state.get("location", ""))
    if r.state.get("fidelity"):
        print(f"Fidelity: {r.state['fidelity']['status']} ({r.state['fidelity']['fails']} fails)")


if __name__ == "__main__":
    main()
