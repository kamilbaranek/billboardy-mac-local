#!/usr/bin/env python3

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from bisect import bisect_left
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import xlrd
from xlutils.copy import copy as xl_copy
from xlwt.Row import StrCell


DEFAULT_ALLOWED_VALUES = ("volný",)
SLASH_PERIOD_RE = re.compile(r"^((?:19|20)\d{2})/(\d{1,2})$")
NAMED_PERIOD_RE = re.compile(r"^((?:19|20)\d{2})\s+(.+)$")
MONTH_NAME_TO_NUMBER = {
    "leden": 1,
    "unor": 2,
    "brezen": 3,
    "duben": 4,
    "kveten": 5,
    "cerven": 6,
    "cervenec": 7,
    "srpen": 8,
    "zari": 9,
    "rijen": 10,
    "listopad": 11,
    "prosinec": 12,
    "january": 1,
    "february": 2,
    "march": 3,
    "april": 4,
    "may": 5,
    "june": 6,
    "july": 7,
    "august": 8,
    "september": 9,
    "october": 10,
    "november": 11,
    "december": 12,
}


class SyncError(RuntimeError):
    pass


@dataclass
class Settings:
    output_root: Path
    mount_names: list[str]
    share_relative_path: str | None
    source_path: Path | None
    allowed_values: set[str]
    header_row_index: int


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Copy a mounted billboard occupancy workbook locally and anonymize booked slots."
    )
    parser.add_argument("--config", type=Path, help="Path to a JSON config file.")
    parser.add_argument("--source", type=Path, help="Explicit source workbook path. Overrides mount lookup.")
    parser.add_argument(
        "--output-root",
        type=Path,
        help="Local directory where raw copies, anonymized copies, and state will be stored.",
    )
    parser.add_argument(
        "--mount-name",
        action="append",
        default=[],
        help="Mounted SMB share name to inspect under /Volumes. Can be passed multiple times.",
    )
    parser.add_argument(
        "--share-relative-path",
        help="Path to the workbook inside the mounted share, for example 'Bookings/rezervace od 2025.xls'.",
    )
    parser.add_argument(
        "--allow-value",
        action="append",
        default=[],
        help="Cell value to keep visible. Defaults to only 'volný'. Can be passed multiple times.",
    )
    parser.add_argument(
        "--header-row-index",
        type=int,
        default=None,
        help="Zero-based header row used for detecting month columns. Defaults to 0.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Inspect and report without writing files.")
    return parser.parse_args()


def load_json_config(path: Path | None) -> dict[str, Any]:
    if path is None:
        return {}
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def build_settings(args: argparse.Namespace, config: dict[str, Any]) -> Settings:
    output_root = Path(
        args.output_root
        or config.get("output_root")
        or Path.cwd() / "output"
    ).expanduser()

    mount_names = args.mount_name or list(config.get("mount_names", []))
    share_relative_path = args.share_relative_path or config.get("share_relative_path")

    source_path_value = args.source or config.get("source_path")
    source_path = Path(source_path_value).expanduser() if source_path_value else None

    allowed_from_config = config.get("allowed_values", list(DEFAULT_ALLOWED_VALUES))
    allowed_values = {value.casefold() for value in (args.allow_value or allowed_from_config)}
    if not allowed_values:
        allowed_values = {value.casefold() for value in DEFAULT_ALLOWED_VALUES}

    header_row_index = (
        args.header_row_index
        if args.header_row_index is not None
        else int(config.get("header_row_index", 0))
    )

    return Settings(
        output_root=output_root,
        mount_names=mount_names,
        share_relative_path=share_relative_path,
        source_path=source_path,
        allowed_values=allowed_values,
        header_row_index=header_row_index,
    )


def resolve_source(settings: Settings) -> Path | None:
    if settings.source_path is not None:
        if not settings.source_path.exists():
            raise SyncError(f"Explicit source file was not found: {settings.source_path}")
        return settings.source_path

    if not settings.mount_names or not settings.share_relative_path:
        raise SyncError("Provide either --source or both mount_names and share_relative_path in the config.")

    volumes_dir = Path("/Volumes")
    if not volumes_dir.exists():
        return None

    for mount_name in settings.mount_names:
        for mount_dir in sorted(volumes_dir.glob(f"{mount_name}*")):
            if not mount_dir.is_dir():
                continue
            candidate = mount_dir / settings.share_relative_path
            if candidate.exists():
                return candidate

    return None


def file_signature(path: Path) -> dict[str, int]:
    stats = path.stat()
    return {"size": stats.st_size, "mtime_ns": stats.st_mtime_ns}


def sha256_for_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def ensure_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def save_json_atomic(path: Path, payload: dict[str, Any]) -> None:
    ensure_directory(path.parent)
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", delete=False, dir=path.parent) as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
        temp_name = handle.name
    os.replace(temp_name, path)


def slugify_filename(name: str) -> str:
    normalized = re.sub(r"[^\w.\-]+", "-", name, flags=re.UNICODE).strip("-")
    return normalized or "workbook"


def timestamp_from_mtime(mtime_ns: int) -> str:
    return datetime.fromtimestamp(mtime_ns / 1_000_000_000).strftime("%Y%m%d-%H%M%S")


def copy_source_to_local_archive(source_path: Path, destination: Path) -> dict[str, int]:
    before = file_signature(source_path)
    ensure_directory(destination.parent)

    with tempfile.NamedTemporaryFile(delete=False, dir=destination.parent, suffix=destination.suffix) as handle:
        temp_path = Path(handle.name)

    try:
        shutil.copy2(source_path, temp_path)
        after = file_signature(source_path)
        if before != after:
            raise SyncError("Source workbook changed during copy. Retry when colleagues are done editing.")
        os.replace(temp_path, destination)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise

    return before


def normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    return "".join(char for char in normalized if not unicodedata.combining(char)).casefold().strip()


def parse_period_header(value: Any) -> tuple[int, int] | None:
    text = str(value).strip()
    if not text:
        return None

    slash_match = SLASH_PERIOD_RE.fullmatch(text)
    if slash_match:
        year = int(slash_match.group(1))
        month = int(slash_match.group(2))
        if 1 <= month <= 12:
            return (year, month)
        return None

    named_match = NAMED_PERIOD_RE.fullmatch(text)
    if not named_match:
        return None

    year = int(named_match.group(1))
    remainder = normalize_text(named_match.group(2))
    month_name = remainder.split()[0]
    month = MONTH_NAME_TO_NUMBER.get(month_name)
    if month is None:
        return None
    return (year, month)


def current_period() -> tuple[int, int]:
    now = datetime.now()
    return (now.year, now.month)


def format_period_for_state(period: tuple[int, int]) -> str:
    return f"{period[0]:04d}-{period[1]:02d}"


def remap_visual_column(old_column: int | None, keep_columns: list[int]) -> int | None:
    if old_column is None or not keep_columns:
        return old_column
    new_column = bisect_left(keep_columns, old_column)
    return min(new_column, len(keep_columns) - 1)


def remap_page_breaks(page_breaks: list[int], keep_columns: list[int]) -> list[int]:
    if not page_breaks or not keep_columns:
        return []
    remapped = {remap_visual_column(page_break, keep_columns) for page_break in page_breaks}
    return sorted(value for value in remapped if value is not None)


def find_last_meaningful_column(sheet: Any) -> int:
    for col_index in range(sheet.ncols - 1, -1, -1):
        for row_index in range(sheet.nrows):
            text = str(sheet.cell_value(row_index, col_index)).strip()
            if text:
                return col_index
    return -1


def replace_cell_preserving_style(workbook: Any, sheet: Any, row_index: int, col_index: int, value: str) -> None:
    row = sheet.row(row_index)
    current_cell = row._Row__cells.get(col_index)
    if current_cell is None:
        raise SyncError(f"Missing copied cell at row {row_index + 1}, col {col_index + 1}.")
    row.insert_cell(col_index, StrCell(row_index, col_index, current_cell.xf_idx, workbook.add_str(value)))


def prune_sheet_columns(sheet: Any, keep_columns: list[int]) -> None:
    column_mapping = {old_col: new_col for new_col, old_col in enumerate(keep_columns)}
    rows = sheet._Worksheet__rows
    cols = sheet._Worksheet__cols

    sheet_min_col: int | None = None
    sheet_max_col: int | None = None

    for row in rows.values():
        old_cells = row._Row__cells
        new_cells = {}
        for old_col, cell in sorted(old_cells.items()):
            new_col = column_mapping.get(old_col)
            if new_col is None:
                continue
            if hasattr(cell, "colx"):
                cell.colx = new_col
            new_cells[new_col] = cell

        row._Row__cells = new_cells
        if new_cells:
            row_columns = sorted(new_cells)
            row._Row__min_col_idx = row_columns[0]
            row._Row__max_col_idx = row_columns[-1]
            sheet_min_col = row_columns[0] if sheet_min_col is None else min(sheet_min_col, row_columns[0])
            sheet_max_col = row_columns[-1] if sheet_max_col is None else max(sheet_max_col, row_columns[-1])
        else:
            row._Row__min_col_idx = 0
            row._Row__max_col_idx = 0

    new_cols = {}
    for old_col, col_obj in sorted(cols.items()):
        new_col = column_mapping.get(old_col)
        if new_col is None:
            continue
        col_obj._index = new_col
        new_cols[new_col] = col_obj
        sheet_min_col = new_col if sheet_min_col is None else min(sheet_min_col, new_col)
        sheet_max_col = new_col if sheet_max_col is None else max(sheet_max_col, new_col)

    sheet._Worksheet__cols = new_cols

    remapped_merged_ranges = []
    for row_low, row_high, col_low, col_high in sheet._Worksheet__merged_ranges:
        kept_range_columns = [column_mapping[col] for col in range(col_low, col_high) if col in column_mapping]
        if len(kept_range_columns) != (col_high - col_low):
            continue
        remapped_merged_ranges.append((row_low, row_high, kept_range_columns[0], kept_range_columns[-1] + 1))
    sheet._Worksheet__merged_ranges = remapped_merged_ranges

    sheet._Worksheet__first_visible_col = remap_visual_column(sheet._Worksheet__first_visible_col, keep_columns) or 0
    sheet._Worksheet__vert_split_first_visible = remap_visual_column(
        sheet._Worksheet__vert_split_first_visible, keep_columns
    )
    sheet._Worksheet__vert_split_pos = remap_visual_column(sheet._Worksheet__vert_split_pos, keep_columns)
    sheet._Worksheet__vert_page_breaks = remap_page_breaks(sheet._Worksheet__vert_page_breaks, keep_columns)

    sheet.first_used_col = 0 if sheet_min_col is None else sheet_min_col
    sheet.last_used_col = 0 if sheet_max_col is None else sheet_max_col


def anonymize_workbook(
    raw_copy_path: Path,
    public_output_path: Path,
    allowed_values: set[str],
    header_row_index: int,
    visible_from_period: tuple[int, int],
) -> dict[str, Any]:
    read_book = xlrd.open_workbook(str(raw_copy_path), formatting_info=True)
    write_book = xl_copy(read_book)

    sheet_stats: list[dict[str, Any]] = []
    total_replacements = 0
    total_visible = 0
    total_removed_columns = 0

    for sheet_index, read_sheet in enumerate(read_book.sheets()):
        write_sheet = write_book.get_sheet(sheet_index)
        if read_sheet.nrows <= header_row_index:
            continue

        period_columns = {
            col: period
            for col in range(read_sheet.ncols)
            if (period := parse_period_header(read_sheet.cell_value(header_row_index, col))) is not None
        }
        if not period_columns:
            continue

        last_meaningful_column = find_last_meaningful_column(read_sheet)
        kept_period_columns = [col for col, period in period_columns.items() if period >= visible_from_period]
        removed_period_columns = [col for col, period in period_columns.items() if period < visible_from_period]
        keep_columns = [
            col
            for col in range(last_meaningful_column + 1)
            if col not in period_columns or col in kept_period_columns
        ]

        replacements = 0
        visible = 0

        for row_index in range(header_row_index + 1, read_sheet.nrows):
            for col_index in kept_period_columns:
                value = read_sheet.cell_value(row_index, col_index)
                text = str(value).strip()
                if not text:
                    continue
                if text.casefold() in allowed_values:
                    visible += 1
                    continue
                replace_cell_preserving_style(write_book, write_sheet, row_index, col_index, "OBSAZENO")
                replacements += 1

        if removed_period_columns:
            prune_sheet_columns(write_sheet, keep_columns)

        total_replacements += replacements
        total_visible += visible
        total_removed_columns += len(removed_period_columns)
        sheet_stats.append(
            {
                "sheet_name": read_sheet.name,
                "period_columns": len(period_columns),
                "kept_period_columns": len(kept_period_columns),
                "removed_past_period_columns": len(removed_period_columns),
                "replacements": replacements,
                "visible_values": visible,
            }
        )

    ensure_directory(public_output_path.parent)
    with tempfile.NamedTemporaryFile(delete=False, dir=public_output_path.parent, suffix=public_output_path.suffix) as handle:
        temp_path = Path(handle.name)

    try:
        write_book.save(str(temp_path))
        os.replace(temp_path, public_output_path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise

    return {
        "sheet_stats": sheet_stats,
        "total_replacements": total_replacements,
        "total_visible_values": total_visible,
        "total_removed_past_period_columns": total_removed_columns,
        "visible_from_period": format_period_for_state(visible_from_period),
    }


def main() -> int:
    args = parse_args()
    config = load_json_config(args.config)
    settings = build_settings(args, config)
    visible_from_period = current_period()

    source_path = resolve_source(settings)
    if source_path is None:
        print("No mounted source workbook found. Nothing to do.")
        return 0

    state_dir = settings.output_root / "state"
    raw_archive_dir = settings.output_root / "raw" / "archive"
    raw_latest_dir = settings.output_root / "raw" / "latest"
    public_archive_dir = settings.output_root / "public" / "archive"
    public_latest_dir = settings.output_root / "public" / "latest"

    state_path = state_dir / "last_sync.json"
    state = load_state(state_path)

    current_signature = file_signature(source_path)
    previous_signature = state.get("source_signature")
    if previous_signature == current_signature and state.get("visible_from_period") == format_period_for_state(visible_from_period):
        print("Source workbook has not changed since the last successful sync.")
        return 0

    stamp = timestamp_from_mtime(current_signature["mtime_ns"])
    raw_archive_name = f"{stamp}__{slugify_filename(source_path.name)}"
    public_archive_name = f"{stamp}__{slugify_filename(source_path.stem)}__anonymized.xls"

    raw_archive_path = raw_archive_dir / raw_archive_name
    raw_latest_path = raw_latest_dir / source_path.name
    public_archive_path = public_archive_dir / public_archive_name
    public_latest_path = public_latest_dir / f"{source_path.stem} - anonymized.xls"

    if args.dry_run:
        print(f"Would sync from: {source_path}")
        print(f"Would archive raw copy to: {raw_archive_path}")
        print(f"Would write anonymized copy to: {public_latest_path}")
        print(f"Would keep periods from: {format_period_for_state(visible_from_period)}")
        return 0

    copy_signature = copy_source_to_local_archive(source_path, raw_archive_path)
    ensure_directory(raw_latest_path.parent)
    shutil.copy2(raw_archive_path, raw_latest_path)

    anonymize_stats = anonymize_workbook(
        raw_copy_path=raw_latest_path,
        public_output_path=public_latest_path,
        allowed_values=settings.allowed_values,
        header_row_index=settings.header_row_index,
        visible_from_period=visible_from_period,
    )
    ensure_directory(public_archive_path.parent)
    shutil.copy2(public_latest_path, public_archive_path)

    public_hash = sha256_for_file(public_latest_path)
    raw_hash = sha256_for_file(raw_latest_path)

    sync_report = {
        "source_path": str(source_path),
        "source_signature": copy_signature,
        "raw_latest_path": str(raw_latest_path),
        "raw_archive_path": str(raw_archive_path),
        "raw_sha256": raw_hash,
        "public_latest_path": str(public_latest_path),
        "public_archive_path": str(public_archive_path),
        "public_sha256": public_hash,
        "synced_at": datetime.now().isoformat(timespec="seconds"),
        "allowed_values": sorted(settings.allowed_values),
        "header_row_index": settings.header_row_index,
        **anonymize_stats,
    }
    save_json_atomic(state_path, sync_report)

    print(f"Synced source: {source_path}")
    print(f"Raw local copy: {raw_latest_path}")
    print(f"Anonymized copy: {public_latest_path}")
    print(
        "Replacements: "
        f"{anonymize_stats['total_replacements']}, kept visible: {anonymize_stats['total_visible_values']}"
    )
    print(f"Removed past period columns: {anonymize_stats['total_removed_past_period_columns']}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SyncError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
