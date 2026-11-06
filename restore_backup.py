"""
Undo a previous bulk_update.py run using the backup file it wrote under
backups/ before making changes. Restores every affected record's fields
back to the values they had before that run.

Usage:
    python restore_backup.py --backup backups/backup_20261106T120000Z.json
    python restore_backup.py --backup backups/backup_20261106T120000Z.json --execute
"""

import argparse
import getpass
import json
import os
import sys
from datetime import datetime, timezone

from sp_client import ConflictError, SharePointClient

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
MAPPING_PATH = os.path.join(BASE_DIR, "field_mapping.json")
AUDIT_LOG_PATH = os.path.join(BASE_DIR, "audit_log.csv")


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def load_mapping():
    with open(MAPPING_PATH, "r", encoding="utf-8") as f:
        fields = json.load(f)
    by_title = {f["Title"]: f for f in fields}
    by_internal = {f["InternalName"]: f for f in fields}
    return by_title, by_internal


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--backup", required=True, help="Path to a backups/backup_*.json file")
    parser.add_argument("--execute", action="store_true", help="Actually write changes (default is dry-run)")
    parser.add_argument("--yes", action="store_true", help="Skip the interactive confirmation prompt")
    args = parser.parse_args()

    with open(args.backup, "r", encoding="utf-8") as f:
        backup_rows = json.load(f)

    config = load_config()
    by_title, by_internal = load_mapping()

    print(f"This will restore {len(backup_rows)} record(s) to their pre-update values:\n")
    for row in backup_rows:
        changes = ", ".join(
            f"{by_internal.get(k, {}).get('Title', k)}: {v!r}" for k, v in row["old_values"].items()
        )
        print(f"  SC ID {row['sc_id']} (item {row['item_id']}): {changes}")

    if not args.execute:
        print("\nDry run only - nothing restored. Re-run with --execute to apply.")
        return

    if not args.yes:
        answer = input(f"\nType CONFIRM to restore these {len(backup_rows)} record(s): ")
        if answer.strip() != "CONFIRM":
            print("Not confirmed - nothing restored.")
            return

    client = SharePointClient(config["site_url"], config["list_title"])
    audit_rows = []
    for row in backup_rows:
        fields_to_set = dict(row["old_values"])  # already internal-name -> value from the backup
        try:
            client.update_item(row["item_id"], fields_to_set)
            print(f"[RESTORED] {row['sc_id']} (item {row['item_id']})")
            for internal_name, value in fields_to_set.items():
                title = by_internal.get(internal_name, {}).get("Title", internal_name)
                audit_rows.append(
                    {
                        "timestamp": datetime.now(timezone.utc).isoformat(),
                        "user": getpass.getuser(),
                        "sc_id": row["sc_id"],
                        "field": title,
                        "old_value": "(reverted)",
                        "new_value": value,
                        "backup_file": f"RESTORE of {args.backup}",
                    }
                )
        except ConflictError as e:
            print(f"[CONFLICT] {row['sc_id']}: {e}")
        except Exception as e:
            print(f"[ERROR] {row['sc_id']}: {e}")

    if audit_rows:
        import pandas as pd

        is_new = not os.path.exists(AUDIT_LOG_PATH)
        pd.DataFrame(audit_rows).to_csv(AUDIT_LOG_PATH, mode="a", index=False, header=is_new)
        print(f"\nAudit trail appended to {AUDIT_LOG_PATH}")


if __name__ == "__main__":
    main()
