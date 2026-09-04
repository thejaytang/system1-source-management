"""Cross-platform workbook and process locking helpers."""

from __future__ import annotations

import json
import os
import re
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterator

try:
    import portalocker
except ImportError:  # The environment doctor reports this before production use.
    portalocker = None


class WorkbookBusyError(RuntimeError):
    pass


def excel_lock_path(workbook: Path) -> Path:
    return workbook.with_name(f"~${workbook.name}")


def excel_appears_open(workbook: Path) -> bool:
    if excel_lock_path(workbook).exists():
        return True
    if os.name != "nt" or not workbook.exists():
        return False
    try:
        import msvcrt

        with workbook.open("r+b") as handle:
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
    except (OSError, PermissionError):
        return True
    return False


def _process_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


@contextmanager
def exclusive_process_lock(lock_path: Path, stale_hours: float = 12) -> Iterator[None]:
    """Hold one cross-platform run lock for the complete Leader cycle."""
    lock_path = lock_path.resolve()
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    if portalocker is not None:
        try:
            with portalocker.Lock(str(lock_path), mode="a+", timeout=0):
                yield
            lock_path.unlink(missing_ok=True)
            return
        except portalocker.exceptions.LockException as exc:
            raise WorkbookBusyError("Another System1 process is already running.") from exc
    if lock_path.exists():
        try:
            payload = json.loads(lock_path.read_text(encoding="utf-8"))
            age_hours = (time.time() - lock_path.stat().st_mtime) / 3600
            stale = age_hours >= stale_hours or not _process_alive(int(payload.get("pid", -1)))
        except (OSError, ValueError, json.JSONDecodeError):
            stale = (time.time() - lock_path.stat().st_mtime) / 3600 >= stale_hours
        if stale:
            lock_path.unlink(missing_ok=True)
        else:
            raise WorkbookBusyError("Another System1 process is already running.")
    descriptor = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o644)
    try:
        os.write(descriptor, json.dumps({"pid": os.getpid(), "started_at": datetime.now().isoformat(timespec="seconds")}).encode("utf-8"))
        os.close(descriptor)
        descriptor = -1
        yield
    finally:
        if descriptor >= 0:
            os.close(descriptor)
        lock_path.unlink(missing_ok=True)


@contextmanager
def system_lock(workbook: Path, stale_hours: float = 12) -> Iterator[None]:
    """Prevent Excel and concurrent System1 processes from writing together."""
    workbook = workbook.resolve()
    if excel_appears_open(workbook):
        raise WorkbookBusyError("The Excel workbook is open. Save and close it before running System1.")
    lock_path = workbook.with_name(".system1-source-management.lock")
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    if portalocker is not None:
        try:
            with portalocker.Lock(str(lock_path), mode="a+", timeout=0):
                yield
            lock_path.unlink(missing_ok=True)
            return
        except portalocker.exceptions.LockException as exc:
            raise WorkbookBusyError("Another System1 process is already running.") from exc

    if lock_path.exists():
        remove_stale = False
        try:
            payload = json.loads(lock_path.read_text(encoding="utf-8"))
            age_hours = (time.time() - lock_path.stat().st_mtime) / 3600
            remove_stale = age_hours >= stale_hours or not _process_alive(int(payload.get("pid", -1)))
        except (OSError, ValueError, json.JSONDecodeError):
            remove_stale = (time.time() - lock_path.stat().st_mtime) / 3600 >= stale_hours
        if remove_stale:
            lock_path.unlink(missing_ok=True)
        else:
            raise WorkbookBusyError("Another System1 process is already running.")
    descriptor = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o644)
    try:
        payload = {"pid": os.getpid(), "started_at": datetime.now().isoformat(timespec="seconds")}
        os.write(descriptor, json.dumps(payload).encode("utf-8"))
        os.close(descriptor)
        descriptor = -1
        yield
    finally:
        if descriptor >= 0:
            os.close(descriptor)
        lock_path.unlink(missing_ok=True)


def validate_windows_path(path: Path, max_length: int = 240) -> list[str]:
    """Return portable-path issues without mutating the path."""
    issues: list[str] = []
    reserved = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
    invalid = '<>:"|?*'
    for index, part in enumerate(path.parts):
        # A Windows drive prefix such as ``C:`` is legal even when this check
        # runs on macOS or Linux during cross-platform validation.
        is_drive_prefix = index == 0 and re.fullmatch(r"[A-Za-z]:", part) is not None
        stem = Path(part).stem.upper()
        if stem in reserved:
            issues.append(f"Windows reserved name: {part}")
        if not is_drive_prefix and any(char in part for char in invalid):
            issues.append(f"Windows-invalid character in path component: {part}")
        if part.endswith((" ", ".")):
            issues.append(f"Windows path component ends with a space or period: {part}")
    if len(str(path)) > max_length:
        issues.append(f"Path is longer than the portable limit of {max_length} characters.")
    return issues
