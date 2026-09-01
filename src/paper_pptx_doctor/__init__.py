"""Verify that the installed ``pptx2`` import belongs to ``python-pptx2``.

Adapted from paper-pptx's distribution doctor: same wheel-integrity
checks (RECORD hashes for every ``pptx2`` package file), retargeted at
the ``python-pptx2`` distribution and its ``__version__`` sentinel.
"""

from __future__ import annotations

import base64
import csv
import hashlib
import hmac
import importlib
import sys
from importlib.metadata import Distribution, PackageNotFoundError, distribution
from io import StringIO
from pathlib import Path, PurePosixPath
from typing import Iterable, Optional, Tuple


class DoctorError(RuntimeError):
    """The installed ``pptx2`` package cannot be trusted as ``python-pptx2``."""


_REMEDY = "python -m pip install --force-reinstall python-pptx2"


def verify_install() -> str:
    """Verify distribution ownership, installed bytes, and the version sentinel.

    Returns the installed ``python-pptx2`` version. Raises :class:`DoctorError`
    without importing ``pptx2`` until its wheel-owned files have been checked.
    """
    dist = _installed_distribution("python-pptx2")
    if dist is None:
        raise DoctorError("python-pptx2 distribution metadata is missing")

    _verify_pptx_record(dist)

    try:
        pptx2 = importlib.import_module("pptx2")
    except Exception as exc:
        raise DoctorError(f"pptx2 cannot be imported: {exc}") from exc

    sentinel = getattr(pptx2, "__version__", None)
    if sentinel is None:
        raise DoctorError("pptx2.__version__ is missing")
    if sentinel != dist.version:
        raise DoctorError(
            "pptx2.__version__ does not match the installed python-pptx2 "
            f"version ({sentinel!r} != {dist.version!r})"
        )
    return dist.version


def main() -> int:
    """Console entry point for ``paper-pptx-doctor``."""
    try:
        version = verify_install()
    except DoctorError as exc:
        print(f"paper-pptx-doctor: FAIL: {exc}", file=sys.stderr)
        print(f"Remedy: {_REMEDY}", file=sys.stderr)
        return 1
    print(f"paper-pptx-doctor: OK (python-pptx2 {version})")
    return 0


def _installed_distribution(name: str) -> Optional[Distribution]:
    try:
        return distribution(name)
    except PackageNotFoundError:
        return None


def _verify_pptx_record(dist: Distribution) -> None:
    record = dist.read_text("RECORD")
    if record is None:
        raise DoctorError("python-pptx2 RECORD is missing")

    entries = tuple(
        (relative_path, hash_spec)
        for relative_path, hash_spec in _pptx_record_entries(record)
        if hash_spec
    )
    if not entries:
        raise DoctorError("python-pptx2 RECORD has no hashed pptx2 package files")

    for relative_path, hash_spec in entries:
        path = Path(dist.locate_file(relative_path))
        if not path.is_file():
            raise DoctorError(f"python-pptx2 file is missing: {relative_path}")
        algorithm, expected = _parse_hash(hash_spec, relative_path)
        actual = _file_digest(path, algorithm)
        if not hmac.compare_digest(actual, expected):
            raise DoctorError(f"python-pptx2 file hash mismatch: {relative_path}")


def _pptx_record_entries(record: str) -> Iterable[Tuple[PurePosixPath, str]]:
    for row in csv.reader(StringIO(record)):
        if len(row) != 3:
            raise DoctorError("python-pptx2 RECORD contains a malformed row")
        raw_path, hash_spec, _size = row
        path = PurePosixPath(raw_path)
        if not path.parts or path.parts[0] != "pptx2":
            continue
        if path.is_absolute() or ".." in path.parts:
            raise DoctorError(f"python-pptx2 RECORD contains an unsafe path: {raw_path}")
        yield path, hash_spec


def _parse_hash(hash_spec: str, relative_path: PurePosixPath) -> Tuple[str, str]:
    try:
        algorithm, expected = hash_spec.split("=", 1)
        hashlib.new(algorithm)
    except (TypeError, ValueError):
        raise DoctorError(
            f"python-pptx2 RECORD has an invalid hash for {relative_path}"
        ) from None
    if not expected:
        raise DoctorError(
            f"python-pptx2 RECORD has an invalid hash for {relative_path}"
        )
    return algorithm, expected.rstrip("=")


def _file_digest(path: Path, algorithm: str) -> str:
    digest = hashlib.new(algorithm)
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return base64.urlsafe_b64encode(digest.digest()).rstrip(b"=").decode("ascii")


__all__ = ["DoctorError", "main", "verify_install"]
