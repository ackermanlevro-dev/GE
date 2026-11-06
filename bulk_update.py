"""
Bulk-update records in the SC Tracker SharePoint list from a CSV or Excel
file. Each row must include the key column (default "SC ID", set in
config.json) plus any other columns you want to update for that SC ID.
Leave a cell blank to leave that field untouched on the matching record.

Column headers in your input file must match the SharePoint *display*
titles exactly (e.g. "Invoice NO", "GRN Date") - run discover_fields.py
first to see the exact titles and confirm spelling/casing.

SAFETY MODEL (this list is shared, business-critical data):
  - Without --execute, this is ALWAYS a dry run: nothing is written.
  - The whole input file is validated before anything is written. If any
    row fails validation (bad SC ID, duplicate match, bad value), the
    entire run aborts with nothing changed - use --allow-partial to
    instead skip just the bad rows and continue with the rest.
  - Before writing, the current value of every field about to change is
    saved to a timestamped backup file under backups/, and a summary is
    printed for you to confirm (type CONFIRM) before anything is sent.
  - Writes use the record's ETag, so if someone else edited a record
    between validation and writing, that row is rejected instead of
    silently overwriting their change.
  - By default only the fields listed in config.json's "editable_fields"
    can be written (Invoice/GRN fields). Pass --allow-all-fields to lift
    that restriction for a one-off need.
  - Every successful write is appended to audit_log.csv (who/when/old/new).

Usage:
    python bulk_update.py --input updates.csv                # dry run
    python bulk_update.py --input updates.csv --execute       # for real
    python bulk_update.py --input updates.xlsx --execute --allow-partial
"""

import argparse
import getpass
import json
import os
import sys
from datetime import datetime, timezone

import pandas as pd

from sp_client import ConflictError, SharePointClient

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
MAPPING_PATH = os.path.join(BASE_DIR, "field_mapping.json")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
AUDIT_LOG_PATH = os.path.join(BASE_DIR, "audit_log.csv")

DATE_TYPES = {"DateTime"}


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def load_mapping():
    if not os.path.exists(MAPPING_PATH):
        sys.exit(
            "field_mapping.json not found. Run `python discover_fields.py` first "
            "to fetch the list's real field names."
        )
    with open(MAPPING_PATH, "r", encoding="utf-8") as f:
        fields = json.load(f)
    return {f["Title"]: f for f in fields}


def read_input(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
    elif ext in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(path, dtype=str)
        df = df.fillna("")
    else:
        sys.exit(f"Unsupported input file type: {ext} (use .csv or .xlsx)")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def coerce_value(raw_value, field_meta):
    value = str(raw_value).strip()
    if value == "":
        return None
    type_name = field_meta["TypeAsString"]
    if type_name in DATE_TYPES:
        parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if pd.isna(parsed):
            raise ValueError(f"Could not parse date value {raw_value!r}")
        return parsed.strftime("%Y-%m-%dT%H:%M:%SZ")
    if type_name in ("Number", "Currency"):
        return float(value)
    return value


def append_audit_log(rows):
    is_new = not os.path.exists(AUDIT_LOG_PATH)
    audit_df = pd.DataFrame(rows)
    audit_df.to_csv(AUDIT_LOG_PATH, mode="a", index=False, header=is_new)


def main():
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--input", required=True, help="Path to .csv or .xlsx file with updates")
    parser.add_argument(
        "--execute", action="store_true", help="Actually write changes (default is dry-run only)"
    )
    parser.add_argument(
        "--allow-partial",
        action="store_true",
        help="If some rows fail validation, skip them and proceed with the rest "
        "(default: abort the whole run with nothing changed)",
    )
    parser.add_argument(
        "--allow-all-fields",
        action="store_true",
        help="Allow updating fields outside config.json's editable_fields whitelist",
    )
    parser.add_argument(
        "--yes", action="store_true", help="Skip the interactive confirmation prompt"
    )
    parser.add_argument(
        "--report",
        default=None,
        help="Path to write a results CSV (default: <input>_results.csv)",
    )
    args = parser.parse_args()

    config = load_config()
    mapping = load_mapping()
    key_column = config["key_column"]
    whitelist = set(config.get("editable_fields", []))

    if key_column not in mapping:
        sys.exit(f"Key column {key_column!r} not found in field_mapping.json")
    key_internal = mapping[key_column]["InternalName"]

    df = read_input(args.input)
    if key_column not in df.columns:
        sys.exit(f"Input file is missing the key column: {key_column!r}")

    update_columns = [c for c in df.columns if c != key_column]
    unknown = [c for c in update_columns if c not in mapping]
    if unknown:
        sys.exit(
            f"These input columns don't match any known SharePoint field title: {unknown}\n"
            "Check spelling against field_mapping.json (run discover_fields.py to refresh it)."
        )
    readonly = [c for c in update_columns if mapping[c]["ReadOnlyField"]]
    if readonly:
        sys.exit(f"These columns are read-only in SharePoint and can't be written to: {readonly}")

    if not args.allow_all_fields and whitelist:
        not_whitelisted = [c for c in update_columns if c not in whitelist]
        if not_whitelisted:
            sys.exit(
                f"These columns aren't in config.json's editable_fields whitelist: {not_whitelisted}\n"
                "Add them to editable_fields if this is intentional, or re-run with --allow-all-fields."
            )

    client = SharePointClient(config["site_url"], config["list_title"])
    select_internal_names = [mapping[c]["InternalName"] for c in update_columns]

    # --- Pass 1: validate every row and resolve its current state, writing nothing yet.
    plan = []  # rows ready to write: {row_num, sc_id, item_id, etag, old_values, new_values}
    problems = []  # rows that failed validation
    skipped = []  # rows with blank key or no changes requested

    for row_num, row in df.iterrows():
        sc_id = str(row[key_column]).strip()
        if sc_id == "":
            skipped.append({"Row": row_num + 2, "SC ID": "", "Status": "SKIPPED", "Detail": "blank key"})
            continue

        try:
            new_values = {}
            for col in update_columns:
                value = coerce_value(row[col], mapping[col])
                if value is not None:
                    new_values[col] = value
        except ValueError as e:
            problems.append({"Row": row_num + 2, "SC ID": sc_id, "Status": "ERROR", "Detail": str(e)})
            continue

        if not new_values:
            skipped.append(
                {"Row": row_num + 2, "SC ID": sc_id, "Status": "SKIPPED", "Detail": "no values to update"}
            )
            continue

        try:
            item = client.find_item_by_key(key_internal, sc_id, select_internal_names)
        except ValueError as e:
            problems.append({"Row": row_num + 2, "SC ID": sc_id, "Status": "ERROR", "Detail": str(e)})
            continue

        if item is None:
            problems.append({"Row": row_num + 2, "SC ID": sc_id, "Status": "NOT FOUND", "Detail": ""})
            continue

        old_values = {col: item.get(mapping[col]["InternalName"]) for col in new_values}
        old_values_internal = {mapping[col]["InternalName"]: item.get(mapping[col]["InternalName"]) for col in new_values}
        plan.append(
            {
                "row": row_num + 2,
                "sc_id": sc_id,
                "item_id": item["Id"],
                "etag": item["__etag"],
                "old_values": old_values,
                "old_values_internal": old_values_internal,
                "new_values": new_values,
            }
        )

    if problems and not args.allow_partial:
        print(f"\n{len(problems)} row(s) failed validation - aborting with nothing changed:\n")
        for p in problems:
            print(f"  Row {p['Row']} (SC ID {p['SC ID']!r}): {p['Status']} - {p['Detail']}")
        print("\nFix these rows and re-run, or pass --allow-partial to skip just these rows.")
        sys.exit(1)

    if not plan:
        print("Nothing to update after validation.")
        report_rows = problems + skipped
        report_df = pd.DataFrame(report_rows)
        report_path = args.report or (os.path.splitext(args.input)[0] + "_results.csv")
        report_df.to_csv(report_path, index=False)
        print(f"Results written to {report_path}")
        return

    # --- Summary + backup, before writing anything.
    print(f"\n{len(plan)} record(s) will be updated"
          f"{' (DRY RUN - nothing will be written)' if not args.execute else ''}:\n")
    for p in plan:
        changes = ", ".join(
            f"{col}: {p['old_values'][col]!r} -> {p['new_values'][col]!r}" for col in p["new_values"]
        )
        print(f"  SC ID {p['sc_id']} (row {p['row']}): {changes}")
    if problems:
        print(f"\n({len(problems)} other row(s) will be skipped due to validation errors - see report)")

    if not args.execute:
        report_rows = (
            [
                {
                    "Row": p["row"],
                    "SC ID": p["sc_id"],
                    "Status": "DRY-RUN",
                    "Detail": json.dumps(p["new_values"]),
                }
                for p in plan
            ]
            + problems
            + skipped
        )
        report_df = pd.DataFrame(report_rows)
        report_path = args.report or (os.path.splitext(args.input)[0] + "_results.csv")
        report_df.to_csv(report_path, index=False)
        print(f"\nDry run only - no changes written. Results written to {report_path}")
        print("Re-run with --execute to apply these changes.")
        return

    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    backup_path = os.path.join(BACKUP_DIR, f"backup_{timestamp}.json")
    with open(backup_path, "w", encoding="utf-8") as f:
        json.dump(
            [
                {"sc_id": p["sc_id"], "item_id": p["item_id"], "old_values": p["old_values_internal"]}
                for p in plan
            ],
            f,
            indent=2,
        )
    print(f"\nBackup of current values written to {backup_path}")

    if not args.yes:
        answer = input(f"\nType CONFIRM to write these {len(plan)} change(s) to SharePoint: ")
        if answer.strip() != "CONFIRM":
            print("Not confirmed - no changes written. Your backup file above is still safe to keep.")
            return

    # --- Pass 2: write.
    results = list(problems) + list(skipped)
    audit_rows = []
    for p in plan:
        try:
            fields_to_set = {mapping[col]["InternalName"]: val for col, val in p["new_values"].items()}
            client.update_item(p["item_id"], fields_to_set, etag=p["etag"])
            results.append(
                {"Row": p["row"], "SC ID": p["sc_id"], "Status": "OK", "Detail": f"item {p['item_id']}"}
            )
            print(f"[OK] {p['sc_id']} (item {p['item_id']})")
            for col in p["new_values"]:
                audit_rows.append(
                    {
                        "timestamp": datetime.now(timezone.utc).isoformat(),
                        "user": getpass.getuser(),
                        "sc_id": p["sc_id"],
                        "field": col,
                        "old_value": p["old_values"][col],
                        "new_value": p["new_values"][col],
                        "backup_file": backup_path,
                    }
                )
        except ConflictError as e:
            results.append({"Row": p["row"], "SC ID": p["sc_id"], "Status": "CONFLICT", "Detail": str(e)})
            print(f"[CONFLICT] {p['sc_id']}: {e}")
        except Exception as e:
            results.append({"Row": p["row"], "SC ID": p["sc_id"], "Status": "ERROR", "Detail": str(e)})
            print(f"[ERROR] {p['sc_id']}: {e}")

    if audit_rows:
        append_audit_log(audit_rows)

    report_df = pd.DataFrame(results)
    report_path = args.report or (os.path.splitext(args.input)[0] + "_results.csv")
    report_df.to_csv(report_path, index=False)

    ok = (report_df["Status"] == "OK").sum()
    errors = report_df["Status"].isin(["ERROR", "CONFLICT"]).sum()
    not_found = (report_df["Status"] == "NOT FOUND").sum()
    print(f"\nDone. {ok} updated, {errors} errors/conflicts, {not_found} not found.")
    print(f"Full results written to {report_path}")
    print(f"Audit trail appended to {AUDIT_LOG_PATH}")


if __name__ == "__main__":
    main()
