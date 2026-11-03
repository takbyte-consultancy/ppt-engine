"""
Command-line Exhibit creation (Agent 2).

  python -m scripts.create_exhibit --asset-type Bond --acronym American_bond_1
  python -m scripts.create_exhibit --asset-type Equity --acronym ABCF --param region=EU --force
"""
import argparse

from exhibit_creation.runner import ExhibitCreationWorkflow
from exhibit_creation.services import default_exhibit_services


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--asset-type", required=True); ap.add_argument("--acronym", required=True)
    ap.add_argument("--param", action="append", default=[], help="datasource param, key=value (repeatable)")
    ap.add_argument("--template-version", type=int); ap.add_argument("--force", action="store_true")
    ap.add_argument("--no-thumbnail", action="store_true")
    a = ap.parse_args()
    params = dict(p.split("=", 1) for p in a.param)
    svc = default_exhibit_services(lambda k, p: print(f"  {p}") if k in ("node", "error") else None,
                                   thumbnails=not a.no_thumbnail)
    r = ExhibitCreationWorkflow(svc).run(a.asset_type, a.acronym, params=params,
                                         template_version=a.template_version, force=a.force)
    print(f"\nResult: {r.status}", r.error or r.state.get("location", ""))
    for w in r.warnings:
        print("  !", w)


if __name__ == "__main__":
    main()
