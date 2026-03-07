#!/usr/bin/env python3
"""Build a deterministic normalized semantic index from .tools/onair_dump."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple


SCHEMA_VERSION = "0.2.0"
DETERMINISTIC_GENERATED_UTC = "1970-01-01T00:00:00Z"

SOURCE_DIRS = ("grammar", "grammar_snapshot", "lookup_index", "runtime", "schema")
KNOWN_LEAGUES = ("mlb", "nba", "nhl")

# Conservative cap for Phase 2 extraction breadth. This keeps the first pass
# high-confidence and diffable while still emitting real normalized objects.
MEASURE_LIMIT_PER_LEAGUE = 10
FILTER_LIMIT_PER_LEAGUE = 8

FILTER_CATEGORY_LOOKUP_KEYS = {
    "playerSplits": "playerSplitsLookup",
    "playerTimes": "playerTimesLookup",
    "teamSplits": "teamSplitsLookup",
    "teamTimes": "teamTimesLookup",
}

ENTITY_TYPE_ALLOWLIST = {
    "DialectEngine_Player",
    "DialectEngine_MlbPlayer",
    "DialectEngine_NbaPlayer",
    "DialectEngine_NhlPlayer",
    "DialectEngine_League",
    "DialectEngine_Coach",
    "DialectEngine_Draft",
    "DialectEngine_Injury",
    "DialectEngine_Venue",
    "Profiles_Profile",
    "Profiles_AvailableGamePlanFilters",
}

SCRIPT_DIR = Path(__file__).resolve().parent
ONAIR_ROOT = SCRIPT_DIR.parent
OUTPUT_PATH = SCRIPT_DIR / "semantic_index.json"

LEAGUE_TOKEN_SPLIT = re.compile(r"[^a-z0-9]+")


def iter_files_sorted(root: Path) -> Iterable[Path]:
    """Yield files under root in deterministic order."""
    files = (path for path in root.rglob("*") if path.is_file())
    return sorted(files, key=lambda p: p.as_posix().lower())


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def safe_dict(value: Any) -> Dict[str, Any]:
    return value if isinstance(value, dict) else {}


def safe_list(value: Any) -> List[Any]:
    return value if isinstance(value, list) else []


def normalize_token(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "unknown"


def display_name(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return "unknown"
    return re.sub(r"\s+", " ", text.replace("_", " "))


def detect_league_from_text(value: Any) -> Optional[str]:
    normalized = normalize_token(value)
    if "mlb" in normalized:
        return "mlb"
    if "nba" in normalized:
        return "nba"
    if "nhl" in normalized:
        return "nhl"
    return None


def make_semantic_id(kind: str, league: Optional[str], token: Any) -> str:
    league_part = normalize_token(league) if league else "global"
    return f"{kind}:{league_part}:{normalize_token(token)}"


def clean_aliases(values: Iterable[Any]) -> List[str]:
    aliases = sorted({str(v).strip() for v in values if str(v).strip()})
    return aliases


def clean_notes(values: Iterable[Any]) -> List[str]:
    notes = sorted({str(v).strip() for v in values if str(v).strip()})
    return notes


def make_record(
    kind: str,
    token: Any,
    name: Any,
    league: Optional[str],
    source_type: str,
    source_path: str,
    source_ref: Any,
    aliases: Iterable[Any] = (),
    notes: Iterable[Any] = (),
) -> Dict[str, Any]:
    record: Dict[str, Any] = {
        "id": make_semantic_id(kind, league, token),
        "name": display_name(name),
        "league": normalize_token(league) if league else None,
        "source_type": source_type,
        "source_path": source_path,
        "source_ref": str(source_ref),
        "aliases": clean_aliases(aliases),
    }
    normalized_notes = clean_notes(notes)
    if normalized_notes:
        record["notes"] = normalized_notes
    return record


def notes_as_list(record: Dict[str, Any]) -> List[str]:
    notes = record.get("notes")
    if isinstance(notes, list):
        return [str(x) for x in notes]
    if isinstance(notes, str):
        return [notes]
    return []


def merge_record(records_by_id: Dict[str, Dict[str, Any]], candidate: Dict[str, Any]) -> None:
    """Normalize duplicates by stable id and preserve provenance notes."""
    record_id = candidate["id"]
    if record_id not in records_by_id:
        records_by_id[record_id] = candidate
        return

    current = records_by_id[record_id]
    current["aliases"] = clean_aliases(current.get("aliases", []) + candidate.get("aliases", []))

    merged_notes = notes_as_list(current) + notes_as_list(candidate)
    merged_notes.append(
        f"merged_from:{candidate['source_type']}:{candidate['source_path']}#{candidate['source_ref']}"
    )
    cleaned_notes = clean_notes(merged_notes)
    if cleaned_notes:
        current["notes"] = cleaned_notes


def records_to_sorted_list(records_by_id: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [records_by_id[key] for key in sorted(records_by_id.keys())]


def discover_sources() -> Tuple[Dict[str, Dict[str, Any]], List[str]]:
    """Inventory stage: discover deterministic raw source files."""
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
    """Inventory stage: infer leagues only from path tokens."""
    found: Set[str] = set()

    for rel_path in file_paths:
        tokens = [t for t in LEAGUE_TOKEN_SPLIT.split(rel_path.lower()) if t]
        for league in KNOWN_LEAGUES:
            if league in tokens:
                found.add(league)

    return sorted(found)


def extract_profiles() -> List[Dict[str, Any]]:
    records: Dict[str, Dict[str, Any]] = {}
    source_path = "runtime/profiles/profiles.response.json"
    payload = load_json(ONAIR_ROOT / source_path)
    profiles = safe_list(safe_dict(safe_dict(payload).get("data")).get("profiles"))

    sortable_profiles: List[Dict[str, Any]] = []
    for item in profiles:
        if isinstance(item, dict):
            sortable_profiles.append(item)

    sortable_profiles.sort(
        key=lambda x: (
            normalize_token(x.get("league")),
            normalize_token(x.get("displayName")),
            normalize_token(x.get("shareId")),
        )
    )

    for profile in sortable_profiles:
        league = normalize_token(profile.get("league")) if profile.get("league") else None
        display = profile.get("displayName") or profile.get("shareId") or "unknown_profile"
        share_id = profile.get("shareId")

        notes: List[str] = []
        if profile.get("season") is not None:
            notes.append(f"season={profile.get('season')}")
        if profile.get("userId"):
            notes.append("user_id_present")

        record = make_record(
            kind="profile",
            token=display,
            name=display,
            league=league,
            source_type="runtime",
            source_path=source_path,
            source_ref=f"profile:{share_id or display}",
            aliases=[share_id] if share_id else [],
            notes=notes,
        )
        merge_record(records, record)

    return records_to_sorted_list(records)


def extract_filters() -> List[Dict[str, Any]]:
    records: Dict[str, Dict[str, Any]] = {}

    # Runtime filter extraction from availableGamePlanFilters responses.
    for league in KNOWN_LEAGUES:
        source_path = f"runtime/leagues/{league}/availableGamePlanFilters.response.json"
        payload = load_json(ONAIR_ROOT / source_path)
        root = safe_dict(safe_dict(payload).get("data")).get("availableGamePlanFilters")
        filter_root = safe_dict(root)

        runtime_candidates: List[Tuple[str, str, Optional[str]]] = []
        for category in sorted(filter_root.keys()):
            items = [x for x in safe_list(filter_root.get(category)) if isinstance(x, dict)]
            for item in items:
                dsl = item.get("dsl")
                if not dsl:
                    continue
                runtime_id = str(item.get("id")) if item.get("id") is not None else None
                runtime_candidates.append((str(dsl), category, runtime_id))

        runtime_candidates.sort(
            key=lambda x: (normalize_token(x[0]), normalize_token(x[1]), normalize_token(x[2]))
        )

        seen_tokens: Set[str] = set()
        for dsl, category, runtime_id in runtime_candidates:
            token = dsl
            if token in seen_tokens:
                continue
            if len(seen_tokens) >= FILTER_LIMIT_PER_LEAGUE:
                break
            seen_tokens.add(token)
            record = make_record(
                kind="filter",
                token=token,
                name=dsl,
                league=league,
                source_type="runtime",
                source_path=source_path,
                source_ref=f"availableGamePlanFilters.{category}.{dsl}",
                aliases=[runtime_id] if runtime_id is not None else [],
                notes=[f"category={category}"],
            )
            merge_record(records, record)

    # Lookup enrichment for filter aliases and provenance.
    for league in KNOWN_LEAGUES:
        source_path = f"lookup_index/{league}/lookup_index.json"
        payload = load_json(ONAIR_ROOT / source_path)
        lookup = safe_dict(payload)

        for category, lookup_key in sorted(FILTER_CATEGORY_LOOKUP_KEYS.items()):
            mapping = safe_dict(lookup.get(lookup_key))
            for raw_key, raw_value in sorted(mapping.items(), key=lambda kv: normalize_token(kv[0])):
                dsl = str(raw_value or raw_key).strip()
                if not dsl:
                    continue
                token = dsl
                target_id = make_semantic_id("filter", league, token)
                if target_id not in records:
                    continue
                aliases = []
                if str(raw_key).strip() != dsl:
                    aliases.append(raw_key)
                record = make_record(
                    kind="filter",
                    token=token,
                    name=dsl,
                    league=league,
                    source_type="lookup_index",
                    source_path=source_path,
                    source_ref=f"{lookup_key}.{raw_key}",
                    aliases=aliases,
                    notes=[f"category={category}"],
                )
                merge_record(records, record)

    return records_to_sorted_list(records)


def extract_measures() -> List[Dict[str, Any]]:
    records: Dict[str, Dict[str, Any]] = {}
    selected_measure_keys: Dict[str, Set[str]] = {league: set() for league in KNOWN_LEAGUES}

    # Runtime extraction is canonical for Phase 2 measure records.
    for league in KNOWN_LEAGUES:
        source_path = f"runtime/leagues/{league}/fetchBaseMeasures.response.json"
        payload = load_json(ONAIR_ROOT / source_path)
        measures = safe_list(safe_dict(safe_dict(payload).get("data")).get("fetchBaseMeasures"))

        sortable_measures = [m for m in measures if isinstance(m, dict)]
        sortable_measures.sort(
            key=lambda x: (
                normalize_token(x.get("onAirCatalogMeasureKey")),
                normalize_token(x.get("queryEngineMeasureKey")),
            )
        )

        for measure in sortable_measures[:MEASURE_LIMIT_PER_LEAGUE]:
            catalog_key = measure.get("onAirCatalogMeasureKey")
            if not catalog_key:
                continue
            selected_measure_keys[league].add(str(catalog_key))
            engine_key = measure.get("queryEngineMeasureKey")
            categories = safe_list(measure.get("categories"))
            notes = [f"format_type={measure.get('formatType') or 'unknown'}"]
            if categories:
                notes.append(f"categories={ '|'.join(str(c) for c in categories) }")

            record = make_record(
                kind="measure",
                token=catalog_key,
                name=catalog_key,
                league=league,
                source_type="runtime",
                source_path=source_path,
                source_ref=f"fetchBaseMeasures.{catalog_key}",
                aliases=[engine_key] if engine_key and engine_key != catalog_key else [],
                notes=notes,
            )
            merge_record(records, record)

    # Lookup measureByCatalogKey enrichment.
    for league in KNOWN_LEAGUES:
        source_path = f"lookup_index/{league}/lookup_index.json"
        payload = load_json(ONAIR_ROOT / source_path)
        lookup = safe_dict(payload)

        catalog_map = safe_dict(lookup.get("measureByCatalogKey"))
        for catalog_key, info in sorted(catalog_map.items(), key=lambda kv: normalize_token(kv[0])):
            if catalog_key not in selected_measure_keys[league]:
                continue
            details = safe_dict(info)
            engine_key = details.get("queryEngineMeasureKey")
            categories = safe_list(details.get("categories"))
            notes = []
            if categories:
                notes.append(f"categories={ '|'.join(str(c) for c in categories) }")
            record = make_record(
                kind="measure",
                token=catalog_key,
                name=details.get("onAirCatalogMeasureKey") or catalog_key,
                league=league,
                source_type="lookup_index",
                source_path=source_path,
                source_ref=f"measureByCatalogKey.{catalog_key}",
                aliases=[engine_key] if engine_key and engine_key != catalog_key else [],
                notes=notes,
            )
            merge_record(records, record)

        # normalizedMeasureLookup alias enrichment for selected catalog keys only.
        normalized_map = safe_dict(lookup.get("normalizedMeasureLookup"))
        for alias_key, info in sorted(normalized_map.items(), key=lambda kv: normalize_token(kv[0])):
            details = safe_dict(info)
            catalog_key = details.get("onAirCatalogMeasureKey")
            if not catalog_key or catalog_key not in selected_measure_keys[league]:
                continue
            aliases = [alias_key, details.get("canonicalValue"), details.get("queryEngineMeasureKey")]
            record = make_record(
                kind="measure",
                token=catalog_key,
                name=catalog_key,
                league=league,
                source_type="lookup_index",
                source_path=source_path,
                source_ref=f"normalizedMeasureLookup.{alias_key}",
                aliases=aliases,
                notes=["lookup_alias_enrichment"],
            )
            merge_record(records, record)

    return records_to_sorted_list(records)


def extract_entities() -> List[Dict[str, Any]]:
    records: Dict[str, Dict[str, Any]] = {}

    # High-confidence entity extraction from schema type list.
    schema_path = "schema/03__schema_types_list.response.json"
    payload = load_json(ONAIR_ROOT / schema_path)
    schema_root = safe_dict(safe_dict(safe_dict(payload).get("data")).get("__schema"))
    types = safe_list(schema_root.get("types"))
    sortable_types = [t for t in types if isinstance(t, dict)]
    sortable_types.sort(key=lambda t: normalize_token(t.get("name")))

    for type_def in sortable_types:
        type_name = type_def.get("name")
        type_kind = type_def.get("kind")
        if type_name not in ENTITY_TYPE_ALLOWLIST:
            continue
        if type_kind not in {"OBJECT", "INTERFACE", "ENUM"}:
            continue
        league = detect_league_from_text(type_name)
        record = make_record(
            kind="entity",
            token=type_name,
            name=type_name,
            league=league,
            source_type="schema",
            source_path=schema_path,
            source_ref=type_name,
            aliases=[],
            notes=[f"kind={type_kind}"],
        )
        merge_record(records, record)

    # Grammar tree currently contains only minimal node count stats in this dump.
    # Hook retained for future entity extraction from richer grammar nodes.
    return records_to_sorted_list(records)


def extract_qualifiers() -> List[Dict[str, Any]]:
    # No strong standalone qualifier object shape is currently present in the dump.
    # Keep deterministic empty output and preserve this hook for Phase 3+.
    return []


def build_index_document() -> dict:
    # Stage 1: deterministic inventory.
    roots, all_files = discover_sources()
    source_roots = sorted([name for name, payload in roots.items() if payload["exists"]])
    leagues = discover_leagues(all_files)

    # Stage 2: deterministic extraction + normalization.
    entities = extract_entities()
    measures = extract_measures()
    filters = extract_filters()
    qualifiers = extract_qualifiers()
    profiles = extract_profiles()

    # Stage 3: stable output document.
    document = {
        "schema_version": SCHEMA_VERSION,
        "generated_utc": DETERMINISTIC_GENERATED_UTC,
        "source_roots": source_roots,
        "leagues": leagues,
        "entities": entities,
        "measures": measures,
        "qualifiers": qualifiers,
        "filters": filters,
        "profiles": profiles,
        "raw_sources": {
            "total_files": len(all_files),
            "roots": roots,
        },
    }
    return document


def write_index(document: dict) -> None:
    """Output stage: write deterministic JSON under .tools/onair_dump/index/ only."""
    serialized = json.dumps(document, indent=2, sort_keys=True)
    OUTPUT_PATH.write_text(serialized + "\n", encoding="utf-8")


def main() -> None:
    index_document = build_index_document()
    write_index(index_document)
    print(f"Wrote semantic index: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
