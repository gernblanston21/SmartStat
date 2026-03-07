#!/usr/bin/env python3
"""Build a deterministic normalized semantic index from .tools/onair_dump."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple


SCHEMA_VERSION = "0.4.0"
DETERMINISTIC_GENERATED_UTC = "1970-01-01T00:00:00Z"

SOURCE_DIRS = ("grammar", "grammar_snapshot", "lookup_index", "runtime", "schema")
KNOWN_LEAGUES = ("mlb", "nba", "nhl")

# Conservative caps keep Phase 3 deterministic and high-confidence.
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

QUERY_PATH_TYPE_BY_KIND = {
    "measure": "measure_entry",
    "filter": "filter_entry",
    "profile": "profile_entry",
    "entity": "entity_entry",
}

MERGED_FROM_PATTERN = re.compile(r"^merged_from:([^:]+):(.+)#(.+)$")
LEAGUE_TOKEN_SPLIT = re.compile(r"[^a-z0-9]+")

SCRIPT_DIR = Path(__file__).resolve().parent
ONAIR_ROOT = SCRIPT_DIR.parent
OUTPUT_PATH = SCRIPT_DIR / "semantic_index.json"


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


def token_parts(value: Any) -> List[str]:
    return [x for x in normalize_token(value).split("_") if x]


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


def detect_league_from_path(path: str) -> Optional[str]:
    tokens = [t for t in LEAGUE_TOKEN_SPLIT.split(path.lower()) if t]
    for league in KNOWN_LEAGUES:
        if league in tokens:
            return league
    return None


def make_semantic_id(kind: str, league: Optional[str], token: Any) -> str:
    league_part = normalize_token(league) if league else "global"
    return f"{kind}:{league_part}:{normalize_token(token)}"


def make_relationship_id(
    rel_type: str,
    league: Optional[str],
    from_id: str,
    to_id: str,
    source_ref: str = "",
) -> str:
    league_part = normalize_token(league) if league else "global"
    parts = [
        "rel",
        normalize_token(rel_type),
        league_part,
        normalize_token(from_id),
        normalize_token(to_id),
    ]
    if source_ref:
        parts.append(normalize_token(source_ref))
    return ":".join(parts)


def make_query_path_id(path_type: str, league: Optional[str], entry_record_id: str) -> str:
    league_part = normalize_token(league) if league else "global"
    return f"path:{normalize_token(path_type)}:{league_part}:{normalize_token(entry_record_id)}"


def clean_aliases(values: Iterable[Any]) -> List[str]:
    return sorted({str(v).strip() for v in values if str(v).strip()})


def clean_notes(values: Iterable[Any]) -> List[str]:
    return sorted({str(v).strip() for v in values if str(v).strip()})


def unique_sorted_strings(values: Iterable[Any]) -> List[str]:
    return sorted({str(v).strip() for v in values if str(v).strip()})


def record_kind_from_id(record_id: str) -> str:
    return record_id.split(":", 1)[0] if ":" in record_id else normalize_token(record_id)


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
    """Deduplicate by stable id and preserve merged provenance notes."""
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
    """Phase 1: deterministic inventory."""
    roots: Dict[str, Dict[str, Any]] = {}
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

    sortable_profiles: List[Dict[str, Any]] = [x for x in profiles if isinstance(x, dict)]
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

    # High-confidence source: runtime availableGamePlanFilters.
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

    # Lookup enrichment only for already confirmed runtime filters.
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

    # High-confidence source: runtime fetchBaseMeasures.
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

    # Lookup enrichment for selected runtime measures only.
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

    # Grammar hook remains intentionally deferred for sparse grammar tree content.
    return records_to_sorted_list(records)


def extract_qualifiers() -> List[Dict[str, Any]]:
    # No strong standalone qualifier object shape currently in dump evidence.
    return []


def split_plain_and_merged_notes(notes: Iterable[str]) -> Tuple[List[str], List[Tuple[str, str, str]]]:
    plain_notes: List[str] = []
    merged_entries: List[Tuple[str, str, str]] = []
    for note in notes:
        match = MERGED_FROM_PATTERN.match(note)
        if not match:
            plain_notes.append(note)
            continue
        merged_entries.append((match.group(1), match.group(2), match.group(3)))
    return sorted(set(plain_notes)), sorted(set(merged_entries))


def build_lineage_entries(record: Dict[str, Any]) -> List[Dict[str, str]]:
    lineage: List[Dict[str, str]] = [
        {
            "role": "primary",
            "source_type": record["source_type"],
            "source_path": record["source_path"],
            "source_ref": record["source_ref"],
        }
    ]

    plain_notes, merged_entries = split_plain_and_merged_notes(notes_as_list(record))
    if plain_notes:
        record["notes"] = plain_notes
    elif "notes" in record:
        del record["notes"]

    for source_type, source_path, source_ref in merged_entries:
        lineage.append(
            {
                "role": "merged",
                "source_type": source_type,
                "source_path": source_path,
                "source_ref": source_ref,
            }
        )

    lineage = sorted(
        lineage,
        key=lambda x: (
            normalize_token(x.get("source_type")),
            normalize_token(x.get("source_path")),
            normalize_token(x.get("source_ref")),
            normalize_token(x.get("role")),
        ),
    )
    return lineage


def lineage_to_evidence(lineage: List[Dict[str, str]]) -> List[str]:
    evidence = [f"{x['source_type']}:{x['source_path']}#{x['source_ref']}" for x in lineage]
    return sorted(set(evidence))


def infer_record_confidence(record: Dict[str, Any], lineage: List[Dict[str, str]]) -> str:
    if len(lineage) > 1:
        return "high"
    if record.get("source_type") == "schema":
        return "medium"
    return "high"


def enrich_records_with_traceability(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    enriched: List[Dict[str, Any]] = []
    for record in records:
        item = dict(record)
        lineage = build_lineage_entries(item)
        item["lineage"] = lineage
        item["evidence"] = lineage_to_evidence(lineage)
        item["confidence"] = infer_record_confidence(item, lineage)
        item["related_ids"] = []
        item["relationship_refs"] = []
        enriched.append(item)
    enriched.sort(key=lambda x: x["id"])
    return enriched


def merge_relationship(
    relationships_by_id: Dict[str, Dict[str, Any]],
    candidate: Dict[str, Any],
) -> None:
    rel_id = candidate["id"]
    if rel_id not in relationships_by_id:
        relationships_by_id[rel_id] = candidate
        return

    current = relationships_by_id[rel_id]
    merged_notes = clean_notes(notes_as_list(current) + notes_as_list(candidate))
    if merged_notes:
        current["notes"] = merged_notes


def make_relationship(
    rel_type: str,
    from_id: str,
    to_id: str,
    league: Optional[str],
    source_type: str,
    source_path: str,
    source_ref: str,
    confidence: str,
    notes: Iterable[str] = (),
) -> Dict[str, Any]:
    rel_id = make_relationship_id(rel_type, league, from_id, to_id, source_ref)
    relation: Dict[str, Any] = {
        "id": rel_id,
        "type": rel_type,
        "from_id": from_id,
        "to_id": to_id,
        "league": league,
        "source_type": source_type,
        "source_path": source_path,
        "source_ref": source_ref,
        "confidence": confidence,
    }
    cleaned_notes = clean_notes(notes)
    if cleaned_notes:
        relation["notes"] = cleaned_notes
    return relation


def generate_profile_to_league_relationships(
    profiles: List[Dict[str, Any]],
    relationships_by_id: Dict[str, Dict[str, Any]],
) -> None:
    for profile in profiles:
        league = profile.get("league")
        if not league:
            continue
        league_node = f"league:{league}"
        rel = make_relationship(
            rel_type="profile_to_league",
            from_id=profile["id"],
            to_id=league_node,
            league=league,
            source_type=profile["source_type"],
            source_path=profile["source_path"],
            source_ref=profile["source_ref"],
            confidence="high",
            notes=["explicit_profile_league_field"],
        )
        merge_relationship(relationships_by_id, rel)


def generate_entity_to_league_relationships(
    entities: List[Dict[str, Any]],
    relationships_by_id: Dict[str, Dict[str, Any]],
) -> None:
    for entity in entities:
        league = entity.get("league")
        if not league:
            continue
        league_node = f"league:{league}"
        rel = make_relationship(
            rel_type="entity_to_league",
            from_id=entity["id"],
            to_id=league_node,
            league=league,
            source_type=entity["source_type"],
            source_path=entity["source_path"],
            source_ref=entity["source_ref"],
            confidence="high",
            notes=["league_inferred_from_entity_name"],
        )
        merge_relationship(relationships_by_id, rel)


def get_record_token_from_id(record_id: str) -> str:
    parts = record_id.split(":", 2)
    if len(parts) < 3:
        return normalize_token(record_id)
    return normalize_token(parts[2])


def generate_measure_to_filter_relationships(
    measures: List[Dict[str, Any]],
    filters: List[Dict[str, Any]],
    relationships_by_id: Dict[str, Dict[str, Any]],
) -> None:
    # Conservative rule only: emit link if filter token appears explicitly in
    # measure id token parts. This avoids speculative graph inflation.
    filters_by_league: Dict[str, List[Dict[str, Any]]] = {league: [] for league in KNOWN_LEAGUES}
    for filt in filters:
        league = filt.get("league")
        if league in filters_by_league:
            filters_by_league[league].append(filt)
    for league in KNOWN_LEAGUES:
        filters_by_league[league].sort(key=lambda x: x["id"])

    for measure in measures:
        league = measure.get("league")
        if league not in filters_by_league:
            continue
        measure_tokens = set(token_parts(get_record_token_from_id(measure["id"])))
        for filt in filters_by_league[league]:
            filter_token = normalize_token(get_record_token_from_id(filt["id"]))
            if not filter_token or filter_token not in measure_tokens:
                continue
            rel = make_relationship(
                rel_type="measure_to_filter",
                from_id=measure["id"],
                to_id=filt["id"],
                league=league,
                source_type=measure["source_type"],
                source_path=measure["source_path"],
                source_ref=f"token_match:{filter_token}",
                confidence="medium",
                notes=["strict_token_overlap_rule"],
            )
            merge_relationship(relationships_by_id, rel)


def generate_record_to_source_relationships(
    all_records: List[Dict[str, Any]],
    relationships_by_id: Dict[str, Dict[str, Any]],
) -> None:
    for record in sorted(all_records, key=lambda x: x["id"]):
        league = record.get("league")
        for lineage in record.get("lineage", []):
            source_node = (
                f"source:{lineage['source_type']}:{normalize_token(lineage['source_path'])}"
            )
            rel = make_relationship(
                rel_type="record_to_source",
                from_id=record["id"],
                to_id=source_node,
                league=league,
                source_type=lineage["source_type"],
                source_path=lineage["source_path"],
                source_ref=lineage["source_ref"],
                confidence="high",
                notes=[f"lineage_role={lineage.get('role', 'unknown')}"],
            )
            merge_relationship(relationships_by_id, rel)


def generate_relationships(
    entities: List[Dict[str, Any]],
    measures: List[Dict[str, Any]],
    filters: List[Dict[str, Any]],
    profiles: List[Dict[str, Any]],
    qualifiers: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    relationships_by_id: Dict[str, Dict[str, Any]] = {}
    all_records = entities + measures + filters + profiles + qualifiers

    generate_profile_to_league_relationships(profiles, relationships_by_id)
    generate_entity_to_league_relationships(entities, relationships_by_id)
    generate_measure_to_filter_relationships(measures, filters, relationships_by_id)
    generate_record_to_source_relationships(all_records, relationships_by_id)

    return [relationships_by_id[key] for key in sorted(relationships_by_id.keys())]


def attach_relationship_refs_to_records(
    records_by_id: Dict[str, Dict[str, Any]],
    relationships: List[Dict[str, Any]],
) -> None:
    relationship_refs: Dict[str, Set[str]] = {record_id: set() for record_id in records_by_id}
    related_ids: Dict[str, Set[str]] = {record_id: set() for record_id in records_by_id}

    for relation in relationships:
        rel_id = relation["id"]
        from_id = relation["from_id"]
        to_id = relation["to_id"]

        if from_id in records_by_id:
            relationship_refs[from_id].add(rel_id)
            related_ids[from_id].add(to_id)
        if to_id in records_by_id:
            relationship_refs[to_id].add(rel_id)
            related_ids[to_id].add(from_id)

    for record_id, record in records_by_id.items():
        record["relationship_refs"] = sorted(relationship_refs[record_id])
        record["related_ids"] = sorted(related_ids[record_id])


def build_trace_index(
    records_by_id: Dict[str, Dict[str, Any]],
) -> Dict[str, Dict[str, Any]]:
    by_record_id: Dict[str, Dict[str, Any]] = {}
    by_source_path_sets: Dict[str, Set[str]] = {}

    for record_id in sorted(records_by_id.keys()):
        record = records_by_id[record_id]
        lineage = record.get("lineage", [])
        by_record_id[record_id] = {
            "lineage": lineage,
            "evidence": record.get("evidence", []),
            "relationship_ids": record.get("relationship_refs", []),
        }
        for lineage_entry in lineage:
            source_path = lineage_entry["source_path"]
            if source_path not in by_source_path_sets:
                by_source_path_sets[source_path] = set()
            by_source_path_sets[source_path].add(record_id)

    by_source_path = {
        source_path: sorted(record_ids)
        for source_path, record_ids in sorted(by_source_path_sets.items())
    }

    return {
        "by_record_id": by_record_id,
        "by_source_path": by_source_path,
    }


def sorted_record_ids_from_relationships(
    entry_record_id: str,
    relationships: List[Dict[str, Any]],
    known_record_ids: Set[str],
) -> List[str]:
    related: Set[str] = set()
    for relation in relationships:
        from_id = relation.get("from_id")
        to_id = relation.get("to_id")
        if from_id == entry_record_id and isinstance(to_id, str) and to_id in known_record_ids:
            related.add(to_id)
        elif to_id == entry_record_id and isinstance(from_id, str) and from_id in known_record_ids:
            related.add(from_id)
    return sorted(related)


def make_query_path_step(
    kind: str,
    ref_id: str,
    label: str,
    source_path: Optional[str],
    source_ref: Optional[str],
) -> Dict[str, Any]:
    return {
        "kind": kind,
        "ref_id": ref_id,
        "label": label,
        "source_path": source_path,
        "source_ref": source_ref,
    }


def generate_query_paths(
    records_by_id: Dict[str, Dict[str, Any]],
    relationships: List[Dict[str, Any]],
    trace_index: Dict[str, Any],
) -> List[Dict[str, Any]]:
    """Phase 4: deterministic query-path generation from high-confidence links."""
    known_record_ids = set(records_by_id.keys())
    trace_by_record_id = safe_dict(trace_index.get("by_record_id"))
    paths: List[Dict[str, Any]] = []

    for record_id in sorted(records_by_id.keys()):
        record = records_by_id[record_id]
        kind = record_kind_from_id(record_id)
        path_type = QUERY_PATH_TYPE_BY_KIND.get(kind)
        if not path_type:
            continue

        league = record.get("league")
        related_record_ids = sorted_record_ids_from_relationships(record_id, relationships, known_record_ids)
        trace_record = safe_dict(trace_by_record_id.get(record_id))
        lineage = [x for x in safe_list(trace_record.get("lineage")) if isinstance(x, dict)]
        lineage.sort(
            key=lambda x: (
                normalize_token(x.get("source_type")),
                normalize_token(x.get("source_path")),
                normalize_token(x.get("source_ref")),
                normalize_token(x.get("role")),
            )
        )

        steps: List[Dict[str, Any]] = [
            make_query_path_step(
                kind="entry_record",
                ref_id=record_id,
                label=record.get("name", record_id),
                source_path=record.get("source_path"),
                source_ref=record.get("source_ref"),
            )
        ]
        if league:
            steps.append(
                make_query_path_step(
                    kind="league_context",
                    ref_id=f"league:{league}",
                    label=f"League {str(league).upper()}",
                    source_path=None,
                    source_ref=None,
                )
            )

        # Keep related-record traversal deterministic and conservative.
        for related_id in related_record_ids:
            related = records_by_id[related_id]
            steps.append(
                make_query_path_step(
                    kind="related_record",
                    ref_id=related_id,
                    label=related.get("name", related_id),
                    source_path=related.get("source_path"),
                    source_ref=related.get("source_ref"),
                )
            )

        for entry in lineage:
            source_type = str(entry.get("source_type") or "unknown")
            source_path = entry.get("source_path")
            source_ref = entry.get("source_ref")
            steps.append(
                make_query_path_step(
                    kind="source_lineage",
                    ref_id=f"source:{source_type}:{normalize_token(source_path)}",
                    label=f"{source_type}:{source_path}",
                    source_path=source_path,
                    source_ref=source_ref,
                )
            )

        source_evidence = unique_sorted_strings(trace_record.get("evidence", []))
        terminal_record_ids = sorted({record_id, *related_record_ids})
        notes = ["deterministic_phase4_query_path"]
        if not league:
            notes.append("no_league_context_step")

        path = {
            "id": make_query_path_id(path_type, league, record_id),
            "league": league,
            "entry_record_id": record_id,
            "path_type": path_type,
            "steps": steps,
            "terminal_record_ids": terminal_record_ids,
            "source_evidence": source_evidence,
            "confidence": record.get("confidence", "medium"),
            "notes": notes,
        }
        paths.append(path)

    paths.sort(key=lambda x: x["id"])
    return paths


def index_records_by_id(records: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    return {record["id"]: record for record in records}


def refresh_records_from_index(
    base_records: List[Dict[str, Any]],
    records_by_id: Dict[str, Dict[str, Any]],
) -> List[Dict[str, Any]]:
    return [records_by_id[record["id"]] for record in base_records]


def infer_misc_bucket(source_type: str, source_path: str) -> str:
    parts = source_path.split("/")
    if len(parts) >= 3 and parts[0] == source_type:
        return parts[1]
    if len(parts) >= 2 and parts[0] == source_type:
        return "root"
    return "misc"


def build_source_tree(
    roots: Dict[str, Dict[str, Any]],
    trace_index: Dict[str, Any],
) -> Dict[str, Any]:
    """Phase 4: deterministic UI-facing source tree for browse/navigation."""
    trace_by_source_path = safe_dict(trace_index.get("by_source_path"))
    root_nodes: List[Dict[str, Any]] = []

    for source_type in SOURCE_DIRS:
        source_payload = safe_dict(roots.get(source_type))
        files = [x for x in safe_list(source_payload.get("files")) if isinstance(x, str)]
        files.sort()

        buckets: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for source_path in files:
            league = detect_league_from_path(source_path)
            if league:
                bucket_node_type = "league_bucket"
                bucket_value = league
                bucket_label = f"League {league.upper()}"
            else:
                bucket_node_type = "misc_bucket"
                bucket_value = infer_misc_bucket(source_type, source_path)
                bucket_label = f"Group {bucket_value}"

            bucket_key = (bucket_node_type, bucket_value)
            bucket = buckets.setdefault(
                bucket_key,
                {
                    "label": bucket_label,
                    "node_type": bucket_node_type,
                    "league": league if league else None,
                    "children": [],
                },
            )

            record_ids = unique_sorted_strings(trace_by_source_path.get(source_path, []))
            label = source_path.split("/", 1)[1] if "/" in source_path else source_path
            leaf_node = {
                "id": f"source_tree:file:{source_type}:{normalize_token(source_path)}",
                "label": label,
                "node_type": "source_file",
                "source_type": source_type,
                "source_path": source_path,
                "league": league if league else None,
                "record_ids": record_ids,
                "children": [],
            }
            bucket["children"].append(leaf_node)

        bucket_nodes: List[Dict[str, Any]] = []
        for bucket_key in sorted(buckets.keys(), key=lambda x: (normalize_token(x[0]), normalize_token(x[1]))):
            bucket = buckets[bucket_key]
            children = sorted(bucket["children"], key=lambda x: x["id"])
            bucket_record_ids = unique_sorted_strings(
                record_id for child in children for record_id in child.get("record_ids", [])
            )
            node = {
                "id": (
                    f"source_tree:bucket:{source_type}:"
                    f"{normalize_token(bucket['node_type'])}:{normalize_token(bucket_key[1])}"
                ),
                "label": bucket["label"],
                "node_type": bucket["node_type"],
                "source_type": source_type,
                "source_path": None,
                "league": bucket["league"],
                "record_ids": bucket_record_ids,
                "children": children,
            }
            bucket_nodes.append(node)

        root_record_ids = unique_sorted_strings(
            record_id for bucket_node in bucket_nodes for record_id in bucket_node.get("record_ids", [])
        )
        root_nodes.append(
            {
                "id": f"source_tree:root:{source_type}",
                "label": source_type,
                "node_type": "source_type_root",
                "source_type": source_type,
                "source_path": None,
                "league": None,
                "record_ids": root_record_ids,
                "children": bucket_nodes,
            }
        )

    return {"roots": root_nodes}


def build_ui_views(records_by_id: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, List[str]]]:
    """Phase 4: compact deterministic helper groupings for future UI."""
    by_league_sets: Dict[str, Set[str]] = {"global": set()}
    for league in KNOWN_LEAGUES:
        by_league_sets[league] = set()

    by_record_type_sets: Dict[str, Set[str]] = {
        "entities": set(),
        "measures": set(),
        "filters": set(),
        "qualifiers": set(),
        "profiles": set(),
    }
    by_source_type_sets: Dict[str, Set[str]] = {source_type: set() for source_type in SOURCE_DIRS}

    for record_id in sorted(records_by_id.keys()):
        record = records_by_id[record_id]
        league = record.get("league")
        by_league_sets[league if league else "global"].add(record_id)

        kind = record_kind_from_id(record_id)
        if kind == "entity":
            by_record_type_sets["entities"].add(record_id)
        elif kind == "measure":
            by_record_type_sets["measures"].add(record_id)
        elif kind == "filter":
            by_record_type_sets["filters"].add(record_id)
        elif kind == "qualifier":
            by_record_type_sets["qualifiers"].add(record_id)
        elif kind == "profile":
            by_record_type_sets["profiles"].add(record_id)

        source_type = str(record.get("source_type") or "")
        if source_type in by_source_type_sets:
            by_source_type_sets[source_type].add(record_id)

    by_league = {key: sorted(values) for key, values in sorted(by_league_sets.items())}
    by_record_type = {key: sorted(values) for key, values in sorted(by_record_type_sets.items())}
    by_source_type = {key: sorted(values) for key, values in sorted(by_source_type_sets.items())}

    return {
        "by_league": by_league,
        "by_record_type": by_record_type,
        "by_source_type": by_source_type,
    }


def build_index_document() -> Dict[str, Any]:
    # Stage 1: deterministic inventory.
    roots, all_files = discover_sources()
    source_roots = sorted([name for name, payload in roots.items() if payload["exists"]])
    leagues = discover_leagues(all_files)

    # Stage 2: deterministic extraction and normalization.
    entities = enrich_records_with_traceability(extract_entities())
    measures = enrich_records_with_traceability(extract_measures())
    filters = enrich_records_with_traceability(extract_filters())
    qualifiers = enrich_records_with_traceability(extract_qualifiers())
    profiles = enrich_records_with_traceability(extract_profiles())

    records_by_id: Dict[str, Dict[str, Any]] = {}
    for record in entities + measures + filters + qualifiers + profiles:
        records_by_id[record["id"]] = record

    # Stage 3: deterministic relationship generation and record linkback.
    relationships = generate_relationships(entities, measures, filters, profiles, qualifiers)
    attach_relationship_refs_to_records(records_by_id, relationships)

    # Refresh category arrays after relationship refs were attached.
    entities = refresh_records_from_index(entities, records_by_id)
    measures = refresh_records_from_index(measures, records_by_id)
    filters = refresh_records_from_index(filters, records_by_id)
    qualifiers = refresh_records_from_index(qualifiers, records_by_id)
    profiles = refresh_records_from_index(profiles, records_by_id)

    # Stage 4: deterministic traceability, query-path, and browse-tree models.
    trace_index = build_trace_index(records_by_id)
    query_paths = generate_query_paths(records_by_id, relationships, trace_index)
    source_tree = build_source_tree(roots, trace_index)
    ui_views = build_ui_views(records_by_id)

    # Stage 5: final document assembly.
    document: Dict[str, Any] = {
        "schema_version": SCHEMA_VERSION,
        "generated_utc": DETERMINISTIC_GENERATED_UTC,
        "source_roots": source_roots,
        "leagues": leagues,
        "entities": entities,
        "measures": measures,
        "qualifiers": qualifiers,
        "filters": filters,
        "profiles": profiles,
        "relationships": relationships,
        "trace_index": trace_index,
        "query_paths": query_paths,
        "source_tree": source_tree,
        "ui_views": ui_views,
        "raw_sources": {
            "total_files": len(all_files),
            "roots": roots,
        },
    }
    return document


def write_index(document: Dict[str, Any]) -> None:
    serialized = json.dumps(document, indent=2, sort_keys=True)
    OUTPUT_PATH.write_text(serialized + "\n", encoding="utf-8")


def main() -> None:
    index_document = build_index_document()
    write_index(index_document)
    print(f"Wrote semantic index: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
