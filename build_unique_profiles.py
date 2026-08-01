import argparse
import glob
import json

import pandas as pd


def parse_dict_field(value):
    if isinstance(value, dict):
        return value
    if isinstance(value, str):
        try:
            return json.loads(value)
        except Exception:
            return {}
    return {}


def extract_agent_slug(hit: dict):
    """
    A listing is posted either by an individual agent (hit['agent_profile'])
    or an agency (hit['agent']). Returns (profile_type, slug) for whichever
    is present, or (None, None).
    """
    agent_profile = parse_dict_field(hit.get("agent_profile"))
    if agent_profile.get("slug"):
        return "agent", agent_profile["slug"]

    agent = parse_dict_field(hit.get("agent"))
    if agent.get("slug"):
        return "agency", agent["slug"]

    return None, None


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--pattern", type=str, required=True, help="glob pattern for raw jsonl files")
    parser.add_argument("--agent-output", type=str, default="combined_agent_profiles.xlsx")
    parser.add_argument("--agency-output", type=str, default="combined_agency_profiles.xlsx")
    args = parser.parse_args()

    files = sorted(glob.glob(args.pattern))
    if not files:
        print(f"No files matched {args.pattern}, nothing to combine.")
        raise SystemExit(0)

    agent_slugs = set()
    agency_slugs = set()

    for path in files:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                hit = json.loads(line)
                profile_type, slug = extract_agent_slug(hit)
                if profile_type == "agent":
                    agent_slugs.add(slug)
                elif profile_type == "agency":
                    agency_slugs.add(slug)

    print(f"Found {len(agent_slugs)} unique agent slugs, {len(agency_slugs)} unique agency slugs "
          f"across {len(files)} raw files.")

    pd.DataFrame({"slug": sorted(agent_slugs)}).to_excel(args.agent_output, index=False)
    pd.DataFrame({"slug": sorted(agency_slugs)}).to_excel(args.agency_output, index=False)

    print(f"Wrote {args.agent_output} and {args.agency_output}")