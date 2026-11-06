"""
Run this once (and again any time columns are added/renamed on the list)
to fetch the SC Tracker list's real field names from SharePoint and save
them to field_mapping.json. bulk_update.py uses that file to translate the
column headers in your input file into what SharePoint actually expects.

Usage:
    python discover_fields.py
"""

import json
import os

from sp_client import SharePointClient

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
MAPPING_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "field_mapping.json")


def main():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)

    client = SharePointClient(config["site_url"], config["list_title"])
    fields = client.get_fields()

    with open(MAPPING_PATH, "w", encoding="utf-8") as f:
        json.dump(fields, f, indent=2)

    print(f"Wrote {len(fields)} fields to {MAPPING_PATH}\n")
    print(f"{'Display Title':35} {'Internal Name':30} Type")
    print("-" * 90)
    for field in sorted(fields, key=lambda x: x["Title"]):
        ro = " (read-only)" if field["ReadOnlyField"] else ""
        print(f"{field['Title']:35} {field['InternalName']:30} {field['TypeAsString']}{ro}")


if __name__ == "__main__":
    main()
