#!/usr/bin/env python3
"""Build a deterministic scaffold semantic index from .tools/onair_dump."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Dict, Iterable, List, Set


SCHEMA_VERSION = "0.1.0"
DETERMINISTIC_GENERATED_UTC = "1970-01-01T00:00:00Z"

SOURCE_DIRS = ("grammar", "grammar_snapshot", "lookup_index", "runtime", "schema")
KNOWN_LEAGUES = ("mlb", "nba", "nhl")

SCRIPT_DIR = Path(__file__).resolve().parent
ONAIR_ROOT = SCRIPT_DIR.parent
OUTPUT_PATH = SCRIPT_DIR / "semantic_index.json"

LEAGUE_TOKEN_SPLIT = re.compile(r"[^a-z0-9]+")


def iter_files_sorted(root: Path) -> Iterable[Path]:
    """Yield files under root in deterministic order."""
    files = (path for path in root.rglob("*") if path.is_file())
    return sorted(files, key=lambda p: p.as_posix().lower())


def discover_sources() -> tuple[Dict[str, dict], List[str]]:
    """Scaffold logic: discover raw source files without semantic inference."""
    roots: Dict[str, dict] = {}
    all_files: List[str] = []

    for source_name in SOURCE_DIRS:
        source_root = ONAIR_ROOT / source_name
        rel_files: List[str] = []

        if source_root.exists():
            for file_path in iter_files_sorted(source_root):
                rel_path = file_path.relative_to(ONAIR_ROOT).as_posix()
                rel_files.append(rel_path)
                all_files.append(rel_path)

        roots[source_name] = {
            "exists": source_root.exists(),
            "file_count": len(rel_files),
            "files": rel_files,
        }

    all_files.sort()
    return roots, all_files


def discover_leagues(file_paths: Iterable[str]) -> List[str]:
    """Scaffold logic: infer leagues only from path tokens."""
    found: Set[str] = set()

    for rel_path in file_paths:
        tokens = [t for t in LEAGUE_TOKEN_SPLIT.split(rel_path.lower()) if t]
        for league in KNOWN_LEAGUES:
            if league in tokens:
                found.add(league)

    return sorted(found)


def build_index_document() -> dict:
    roots, all_files = discover_sources()
    source_roots = sorted([name for name, payload in roots.items() if payload["exists"]])
    leagues = discover_leagues(all_files)

    # TODO: Replace placeholders with extracted semantic objects in later phases.
    document = {
        "schema_version": SCHEMA_VERSION,
        "generated_utc": DETERMINISTIC_GENERATED_UTC,
        "source_roots": source_roots,
        "leagues": leagues,
        "entities": [],
        "measures": [],
        "qualifiers": [],
        "filters": [],
        "profiles": [],
        "raw_sources": {
            "total_files": len(all_files),
            "roots": roots,
        },
    }
    return document


def write_index(document: dict) -> None:
    """Write deterministic JSON output under .tools/onair_dump/index/ only."""
    serialized = json.dumps(document, indent=2, sort_keys=True)
    OUTPUT_PATH.write_text(serialized + "\n", encoding="utf-8")


def main() -> None:
    index_document = build_index_document()
    write_index(index_document)
    print(f"Wrote semantic index: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
