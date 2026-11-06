# SC Tracker bulk updater

Bulk-update records in the "Asset MGMT SC Tracker" SharePoint list
(Invoice NO, GRN NO, etc.) from a CSV or Excel file, instead of editing
them one by one in the SharePoint UI.

This list is shared, business-critical data, so the tool is built to be
hard to misuse by accident:

- **Dry-run by default.** Nothing is ever written unless you pass `--execute`.
- **Whole-file validation before any write.** If any row has a bad SC ID,
  matches more than one record, or has a value that won't parse, the
  entire run aborts with nothing changed (unless you pass `--allow-partial`).
- **Automatic backup before writing.** The current value of every field
  about to change is saved to `backups/backup_<timestamp>.json` first.
  `restore_backup.py` can undo a run from that file.
- **Concurrency-safe writes.** Each write checks the record hasn't been
  edited by someone else since it was read; if it has, that row is
  rejected rather than silently overwritten.
- **Field whitelist.** Only the columns listed in `config.json`'s
  `editable_fields` (Invoice/GRN fields by default) can be written to,
  unless you explicitly pass `--allow-all-fields`.
- **Confirmation + audit trail.** You're shown exactly what will change
  and must type `CONFIRM`; every successful write is appended to
  `audit_log.csv` (who, when, old value, new value).

## Setup

```bash
pip install -r requirements.txt
```

No Azure AD app registration is required. Authentication uses device-code
sign-in via a Microsoft first-party public client app (the same mechanism
PnP PowerShell's `-DeviceLogin` uses), so the first time any script runs
you'll see something like:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code ABCD-EFGH to authenticate.
```

Open that link in any browser, sign in with your GE account, and the
script continues automatically. A token is cached in `.msal_cache.json`
(contains a refresh token - keep it private, don't commit or share it) so
you won't need to do this every run.

If sign-in fails with an error mentioning admin consent (AADSTS65001 /
AADSTS70011), edit `config.json` and change `client_id` to
`31359c7f-bd7e-475c-86db-fdb8c937548e` (PnP Management Shell) and try
again. If that also fails, your tenant is blocking both standard
Microsoft automation apps and you'll need IT to register a dedicated app
with Sites.Selected permission on this site.

## One-time: discover the list's real field names

SharePoint internal field names don't always match what's displayed in
the UI. Run this once (and again if columns are ever added/renamed):

```bash
python discover_fields.py
```

This writes `field_mapping.json` and prints a table of display titles
and SharePoint internal names. `bulk_update.py` uses this file to
translate your input file's column headers.

## Bulk-updating records

1. Build a CSV or Excel file with a column for `SC ID` (must match
   exactly - see `config.json`'s `key_column`) plus any columns you want
   to update, using the same column headers as the SharePoint UI (see
   `sample_input.csv` for an example). Leave a cell blank to leave that
   field untouched for that row.

2. Preview what would change (this is also what happens if you forget
   `--execute` - it's always safe to just run this):

   ```bash
   python bulk_update.py --input updates.csv
   ```

3. Review the printed summary, then run it for real. You'll see a backup
   file path and be asked to type `CONFIRM`:

   ```bash
   python bulk_update.py --input updates.csv --execute
   ```

Each run writes a `<input>_results.csv` report next to your input file,
with a Status column (`OK`, `NOT FOUND`, `ERROR`, `CONFLICT`, `SKIPPED`)
and detail per row, so you can see exactly what happened to each SC ID.

## Undoing a run

Every `--execute` run writes a backup file under `backups/` before
changing anything. To revert those exact changes:

```bash
python restore_backup.py --backup backups/backup_20261106T120000Z.json          # preview
python restore_backup.py --backup backups/backup_20261106T120000Z.json --execute # apply
```

## Audit trail

`audit_log.csv` accumulates one row per field changed, across every run:
timestamp, user, SC ID, field, old value, new value, and which backup
file it came from - useful for tracing "who changed this and when" later.

## Notes

- Dates in the input file can be in `dd/mm/yyyy` or `yyyy-mm-dd` format.
- Matching is by exact value on the key column (`SC ID` by default); if
  more than one record shares the same SC ID, that row is reported as an
  error rather than guessed at.
- `config.json` controls the site URL, list title, key column, auth
  client ID, and the `editable_fields` whitelist if you ever need to
  change any of them.
- `backups/` and `audit_log.csv` contain historical field values -
  don't delete them without checking if you still need them.
