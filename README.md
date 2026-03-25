# Billboard Occupancy Sync

This utility copies the source `.xls` workbook from a mounted SMB share to your Mac and creates a local anonymized copy.

Important guarantee: the network source workbook is never opened for writing. The flow is always:

1. Read-only copy from SMB to a local raw archive.
2. Duplicate that local raw archive into a local `raw/latest` copy.
3. Build the anonymized workbook from the local copy only.

## What gets anonymized

- Any non-empty cell under a month-like header becomes `OBSAZENO`.
- Only exact values listed in `allowed_values` stay visible.
- The default keeps only `volný`.
- Month columns older than the current month are removed from the anonymized output.
- The raw local copy always stays untouched and keeps the full original workbook.

Month-like headers are auto-detected across all sheets, for example:

- `2025/1`
- `2026/12`
- `2005 listopad`

This means the script anonymizes both the current grid in `List1` and any smaller month grids like `List2`, while keeping panel IDs, city names, and other non-period columns intact.

## Files

- [sync_billboard_occupancy.py](/Users/kamilbaranek/dev/billboardy/sync_billboard_occupancy.py): sync and anonymization logic
- [run_sync.sh](/Users/kamilbaranek/dev/billboardy/run_sync.sh): non-interactive launcher for Terminal or `launchd`
- [run_sync.command](/Users/kamilbaranek/dev/billboardy/run_sync.command): double-clickable launcher
- [build_sync_app.sh](/Users/kamilbaranek/dev/billboardy/build_sync_app.sh): builds a Finder-launchable `.app`
- [config.json](/Users/kamilbaranek/dev/billboardy/config.json): local config
- [launchd/com.billboardy.occupancy-sync.plist.example](/Users/kamilbaranek/dev/billboardy/launchd/com.billboardy.occupancy-sync.plist.example): mount-trigger example

## First run

1. Edit [config.json](/Users/kamilbaranek/dev/billboardy/config.json).
2. Set `share_relative_path` to the path of the workbook inside `/Volumes/COMPANY`.
3. Run:

```bash
/Users/kamilbaranek/dev/billboardy/run_sync.sh
```

If the share is not mounted, the script exits immediately with `Nothing to do.`

## Double-click on macOS

You have two options:

1. Double-click [run_sync.command](/Users/kamilbaranek/dev/billboardy/run_sync.command) in Finder.
2. Build a small app bundle and then double-click the app:

```bash
/Users/kamilbaranek/dev/billboardy/build_sync_app.sh
```

That creates `dist/Billboardy Sync.app`, which runs the sync in the background and then shows a result dialog.

## Output layout

The script creates:

- `output/raw/archive/`: versioned local copies of the original `.xls`
- `output/raw/latest/`: the latest local raw copy
- `output/public/archive/`: versioned anonymized workbooks
- `output/public/latest/`: the latest anonymized workbook
- `output/state/last_sync.json`: last successful sync metadata, hashes, and replacement counts

## Manual testing without SMB

You can point the script directly at a local workbook:

```bash
/Users/kamilbaranek/dev/billboardy/run_sync.sh --source "/Users/kamilbaranek/Downloads/rezervace od 2025.xls"
```

## Mount-triggered run on macOS

1. Copy [launchd/com.billboardy.occupancy-sync.plist.example](/Users/kamilbaranek/dev/billboardy/launchd/com.billboardy.occupancy-sync.plist.example) to `~/Library/LaunchAgents/com.billboardy.occupancy-sync.plist`.
2. Adjust the hardcoded paths if you move this folder.
3. Load it:

```bash
launchctl unload ~/Library/LaunchAgents/com.billboardy.occupancy-sync.plist 2>/dev/null || true
launchctl load ~/Library/LaunchAgents/com.billboardy.occupancy-sync.plist
```

With `StartOnMount`, the job wakes up when a filesystem is mounted. The script then checks whether `/Volumes/COMPANY...` contains the configured workbook. If not, it exits cleanly.

## Future Google Sheets step

The script is already split logically into:

- source resolution
- raw local archival
- anonymization

The next step can reuse the same local raw copy and push either:

- the anonymized workbook, or
- a normalized row/column export

to Google Sheets without ever writing back to the SMB source.
