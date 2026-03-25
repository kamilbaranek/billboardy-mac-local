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
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import xlrd
from xlutils.copy import copy as xl_copy
from xlwt.Row import StrCell


PERIOD_HEADER_RE = re.compile(r"^(?:19|20)\d{2}(?:/\d{1,2}|(?:\s+[^\W\d_]+)+)$", re.UNICODE)
DEFAULT_ALLOWED_VALUES = ("volný",)


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


def is_period_header(value: Any) -> bool:
    text = str(value).strip()
    return bool(PERIOD_HEADER_RE.fullmatch(text))


def replace_cell_preserving_style(workbook: Any, sheet: Any, row_index: int, col_index: int, value: str) -> None:
    row = sheet.row(row_index)
    current_cell = row._Row__cells.get(col_index)
    if current_cell is None:
        raise SyncError(f"Missing copied cell at row {row_index + 1}, col {col_index + 1}.")
    row.insert_cell(col_index, StrCell(row_index, col_index, current_cell.xf_idx, workbook.add_str(value)))


def anonymize_workbook(raw_copy_path: Path, public_output_path: Path, allowed_values: set[str], header_row_index: int) -> dict[str, Any]:
    read_book = xlrd.open_workbook(str(raw_copy_path), formatting_info=True)
    write_book = xl_copy(read_book)

    sheet_stats: list[dict[str, Any]] = []
    total_replacements = 0
    total_visible = 0

    for sheet_index, read_sheet in enumerate(read_book.sheets()):
        write_sheet = write_book.get_sheet(sheet_index)
        if read_sheet.nrows <= header_row_index:
            continue

        target_columns = [col for col in range(read_sheet.ncols) if is_period_header(read_sheet.cell_value(header_row_index, col))]
        if not target_columns:
            continue

        replacements = 0
        visible = 0

        for row_index in range(header_row_index + 1, read_sheet.nrows):
            for col_index in target_columns:
                value = read_sheet.cell_value(row_index, col_index)
                text = str(value).strip()
                if not text:
                    continue
                if text.casefold() in allowed_values:
                    visible += 1
                    continue
                replace_cell_preserving_style(write_book, write_sheet, row_index, col_index, "OBSAZENO")
                replacements += 1

        total_replacements += replacements
        total_visible += visible
        sheet_stats.append(
            {
                "sheet_name": read_sheet.name,
                "period_columns": len(target_columns),
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
    }


def main() -> int:
    args = parse_args()
    config = load_json_config(args.config)
    settings = build_settings(args, config)

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
    if previous_signature == current_signature:
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
        return 0

    copy_signature = copy_source_to_local_archive(source_path, raw_archive_path)
    ensure_directory(raw_latest_path.parent)
    shutil.copy2(raw_archive_path, raw_latest_path)

    anonymize_stats = anonymize_workbook(
        raw_copy_path=raw_latest_path,
        public_output_path=public_latest_path,
        allowed_values=settings.allowed_values,
        header_row_index=settings.header_row_index,
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
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SyncError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
