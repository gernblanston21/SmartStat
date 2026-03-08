#!/usr/bin/env python3
"""
Phase 0 OnAir semantic reference extraction (v3 workbook sources).
Deterministic, standard-library-only extraction from XLSX internals.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple
from xml.etree import ElementTree as ET

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

ROOT = Path(__file__).resolve().parents[2]
DOCS_DIR = ROOT / "docs" / "onair"
SOURCE_DIR = DOCS_DIR / "source"

SOURCES = {
    "MLB": SOURCE_DIR / "MLB_OnAir_v3_Stat_Syntax_v3.xlsx",
    "NHL": SOURCE_DIR / "NHL_OnAir_v3_Stat_Syntax_v3.xlsx",
}


def clean(value: str | None) -> str:
    return (value or "").strip().replace("\n", " ").replace("\r", " ")


def header_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", clean(value).lower()).strip("_")


def escape_md(value: str) -> str:
    return value.replace("|", "\\|")


def join_nonempty(values: Iterable[str], sep: str = "; ") -> str:
    return sep.join(clean(v) for v in values if clean(v))


def markdown_table(headers: Sequence[str], rows: Sequence[Dict[str, str]]) -> str:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(escape_md(str(row.get(h, ""))) for h in headers) + " |")
    return "\n".join(lines)


@dataclass
class Sheet:
    name: str
    order: int
    path: str
    headers: List[str]
    rows: List[Dict[str, str]]
    header_row_index: int


@dataclass
class Workbook:
    league: str
    file_name: str
    sheets: List[Sheet]


def read_rows(sheet_xml: ET.Element, shared_strings: List[str]) -> List[List[str]]:
    rows: List[List[str]] = []
    for row in sheet_xml.findall("m:sheetData/m:row", NS):
        cells: List[str] = []
        for cell in row.findall("m:c", NS):
            ctype = cell.attrib.get("t")
            value_node = cell.find("m:v", NS)
            if ctype == "inlineStr":
                cells.append(clean("".join(t.text or "" for t in cell.findall(".//m:t", NS))))
                continue
            if value_node is None:
                cells.append("")
                continue
            raw = value_node.text or ""
            if ctype == "s":
                try:
                    cells.append(clean(shared_strings[int(raw)]))
                except Exception:
                    cells.append(clean(raw))
            else:
                cells.append(clean(raw))
        rows.append(cells)
    return rows


def detect_header(rows: List[List[str]]) -> Tuple[int, List[str]]:
    for idx, row in enumerate(rows):
        if any(clean(v) for v in row):
            return idx, [clean(v) for v in row]
    return 0, []


def build_records(rows: List[List[str]], header_idx: int, headers: List[str]) -> List[Dict[str, str]]:
    width = max([len(headers)] + [len(r) for r in rows] + [0])
    padded_headers = headers + [""] * (width - len(headers))
    records: List[Dict[str, str]] = []
    for row in rows[header_idx + 1 :]:
        vals = row + [""] * (width - len(row))
        if not any(clean(v) for v in vals):
            continue
        rec: Dict[str, str] = {}
        for i, value in enumerate(vals):
            key = padded_headers[i] if padded_headers[i] else f"column_{i + 1}"
            rec[key] = clean(value)
        records.append(rec)
    return records


def load_workbook(path: Path, league: str) -> Workbook:
    with zipfile.ZipFile(path) as zf:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

        rel_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_xml.findall("pr:Relationship", NS)
        }

        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst.findall("m:si", NS):
                shared_strings.append(clean("".join(t.text or "" for t in si.findall(".//m:t", NS))))

        sheets: List[Sheet] = []
        for order, sheet_node in enumerate(wb_xml.findall("m:sheets/m:sheet", NS), start=1):
            name = sheet_node.attrib["name"]
            rid = sheet_node.attrib[
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            ]
            target = rel_map[rid]
            if not target.startswith("xl/"):
                target = f"xl/{target}"

            sheet_xml = ET.fromstring(zf.read(target))
            rows = read_rows(sheet_xml, shared_strings)
            header_idx, headers = detect_header(rows)
            records = build_records(rows, header_idx, headers)

            sheets.append(
                Sheet(
                    name=name,
                    order=order,
                    path=target,
                    headers=headers,
                    rows=records,
                    header_row_index=header_idx + 1,
                )
            )

    return Workbook(league=league, file_name=path.name, sheets=sheets)


def find_sheet(workbook: Workbook, exact_name: str) -> Sheet | None:
    name_l = exact_name.lower()
    for sheet in workbook.sheets:
        if sheet.name.lower() == name_l:
            return sheet
    return None


def row_keys(row: Dict[str, str]) -> Dict[str, str]:
    return {header_key(k): v for k, v in row.items()}


def split_csv_tokens(value: str) -> List[str]:
    # Deterministic split that avoids splitting commas inside parentheses.
    txt = clean(value)
    if not txt:
        return []
    parts: List[str] = []
    buff: List[str] = []
    depth = 0
    for ch in txt:
        if ch == "(":
            depth += 1
        elif ch == ")" and depth > 0:
            depth -= 1
        if ch == "," and depth == 0:
            token = clean("".join(buff))
            if token:
                parts.append(token)
            buff = []
        else:
            buff.append(ch)
    token = clean("".join(buff))
    if token:
        parts.append(token)
    return parts


def parse_filter_and_parameter(syntax: str) -> Tuple[str, str]:
    syntax = clean(syntax)
    m = re.match(r"^([^(]+)\(([^)]+)\)$", syntax)
    if not m:
        return syntax, ""
    return clean(m.group(1)), clean(m.group(2))


def classify_parameter_type(filter_name: str, parameter_examples: str) -> str:
    f = clean(filter_name).lower()
    p = clean(parameter_examples).lower()

    if not p:
        return "none"
    if "tricode" in p:
        return "team_tricode"
    if "[year]" in p or re.search(r"\byear\b", p):
        return "year"
    if re.search(r"\b\d{1,2}:\d{2}\b", p) or "#:##" in p or "clock" in f:
        return "clock_time"
    if "date" in f or "date" in p:
        return "date"
    if "margin" in f or "#-#" in p and any(op in p for op in ["<", ">", "="]):
        return "score_margin"
    if any(op in p for op in ["<#", ">#", "=>#", "<=#", "=#", "<", ">", "="]):
        return "comparison"
    if re.search(r"\b#\b", p):
        return "integer"
    if any(word in p for word in ["player", "team", "coach", "entity"]):
        return "entity_reference"
    if "-" in p or "through" in p or " to " in p:
        return "range"
    return "unknown"


def sorted_unique_dict_rows(rows: List[Dict[str, str]], key_fields: Sequence[str]) -> List[Dict[str, str]]:
    dedup: Dict[Tuple[str, ...], Dict[str, str]] = {}
    for row in rows:
        key = tuple(clean(row.get(f, "")) for f in key_fields)
        if key not in dedup:
            dedup[key] = row
    return [dedup[k] for k in sorted(dedup.keys())]

def extract_measures(workbooks: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    rows: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []

    mapping = [
        ("MLB", "CATEGORY_TO_MEASURE", "general"),
        ("MLB", "CATEGORY_TO_MEASURE_PITCHER", "pitcher"),
        ("MLB", "Additional Measures", "additional"),
        ("NHL", "CATEGORY_TO_MEASURE", "general"),
        ("NHL", "CATEGORY_TO_MEASURE_GOALIE", "goalie"),
        ("NHL", "Additional Measures", "additional"),
    ]

    for league, sheet_name, subtype in mapping:
        sheet = find_sheet(workbooks[league], sheet_name)
        if sheet is None:
            notes.append(f"Missing expected sheet `{sheet_name}` in `{workbooks[league].file_name}`.")
            continue
        used.append(sheet)
        for row in sheet.rows:
            keys = row_keys(row)
            measure_name = clean(keys.get("measure_syntax") or keys.get("measure") or "")
            if not measure_name:
                continue
            rows.append(
                {
                    "measure_name": measure_name,
                    "description": clean(keys.get("description", "")),
                    "league": league,
                    "subtype": subtype,
                    "source_sheet": sheet.name,
                }
            )

    rows = sorted_unique_dict_rows(rows, ("league", "subtype", "measure_name", "source_sheet"))
    return rows, notes, used


def extract_filters(workbooks: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    rows: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []

    for league in ("MLB", "NHL"):
        sheet = find_sheet(workbooks[league], "QUALIFIER_TO_FILTER")
        if sheet is None:
            notes.append(f"Missing expected sheet `QUALIFIER_TO_FILTER` in `{workbooks[league].file_name}`.")
            continue
        used.append(sheet)

        for row in sheet.rows:
            keys = row_keys(row)
            syntax = clean(keys.get("filter_qualifier_syntax") or keys.get("qualifier_filter") or "")
            params = clean(keys.get("available_parameters", ""))
            examples = clean(keys.get("example_s", ""))
            if not syntax:
                continue
            for token in split_csv_tokens(syntax):
                filter_name, inline_param = parse_filter_and_parameter(token)
                merged_examples = join_nonempty([params, examples], " | ")
                parameter_type = classify_parameter_type(filter_name, params or inline_param)
                rows.append(
                    {
                        "filter_name": filter_name,
                        "parameter_type": parameter_type,
                        "parameter_examples": merged_examples,
                        "league": league,
                        "notes": "inline parameter form" if inline_param and not params else "",
                    }
                )

    rows = sorted_unique_dict_rows(rows, ("league", "filter_name", "parameter_examples"))
    return rows, notes, used


def extract_alias_map(workbooks: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    rows: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []

    for league in ("MLB", "NHL"):
        sheet = find_sheet(workbooks[league], "QUALIFIER_TO_FILTER")
        if sheet is None:
            continue
        used.append(sheet)
        for row in sheet.rows:
            keys = row_keys(row)
            syntax = clean(keys.get("filter_qualifier_syntax") or keys.get("qualifier_filter") or "")
            params = clean(keys.get("available_parameters", ""))
            examples = clean(keys.get("example_s", ""))
            if not syntax:
                continue
            syntax_lower = syntax.lower()
            tokens = [parse_filter_and_parameter(token)[0].lower() for token in split_csv_tokens(syntax)]

            # Conservative inferred alias phrases from explicit parameterized filter forms.
            if "all_star_break" in tokens and "before" in (params + " " + syntax_lower):
                rows.append(
                    {
                        "alias_phrase": "PRE ALL-STAR BREAK",
                        "canonical_filter": "all_star_break",
                        "canonical_parameter": "before",
                        "league": league,
                        "example": examples,
                        "notes": "Inferred from explicit before/after parameter semantics.",
                    }
                )
            if "all_star_break" in tokens and "after" in (params + " " + syntax_lower):
                rows.append(
                    {
                        "alias_phrase": "POST ALL-STAR BREAK",
                        "canonical_filter": "all_star_break",
                        "canonical_parameter": "after",
                        "league": league,
                        "example": examples,
                        "notes": "Inferred from explicit before/after parameter semantics.",
                    }
                )
                rows.append(
                    {
                        "alias_phrase": "SINCE ALL-STAR BREAK",
                        "canonical_filter": "all_star_break",
                        "canonical_parameter": "after",
                        "league": league,
                        "example": examples,
                        "notes": "Inferred from explicit before/after parameter semantics.",
                    }
                )

            if "last_game" in tokens:
                rows.append(
                    {
                        "alias_phrase": "LAST N GAMES",
                        "canonical_filter": "last_game",
                        "canonical_parameter": "N",
                        "league": league,
                        "example": examples,
                        "notes": "Inferred from numeric placeholder examples.",
                    }
                )

            if "season" in tokens and "last#" in params.lower():
                rows.append(
                    {
                        "alias_phrase": "LAST N SEASONS",
                        "canonical_filter": "season",
                        "canonical_parameter": "lastN",
                        "league": league,
                        "example": examples,
                        "notes": "Inferred from `last#` parameter pattern.",
                    }
                )
            if "season" in tokens and "prev#" in params.lower():
                rows.append(
                    {
                        "alias_phrase": "PREVIOUS N SEASONS",
                        "canonical_filter": "season",
                        "canonical_parameter": "prevN",
                        "league": league,
                        "example": examples,
                        "notes": "Inferred from `prev#` parameter pattern.",
                    }
                )

    if not rows:
        notes.append("No confident alias mappings were derivable from current sheet structure.")

    rows = sorted_unique_dict_rows(rows, ("league", "alias_phrase", "canonical_filter", "canonical_parameter"))
    return rows, notes, used


def extract_entities(workbooks: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    rows: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []

    for league in ("MLB", "NHL"):
        sheet = find_sheet(workbooks[league], "Entities")
        if sheet is None:
            notes.append(f"Missing expected sheet `Entities` in `{workbooks[league].file_name}`.")
            continue
        used.append(sheet)
        for row in sheet.rows:
            keys = row_keys(row)
            entity_name = clean(keys.get("entity_syntax") or keys.get("entity") or "")
            params = clean(keys.get("available_parameters", ""))
            if not entity_name:
                continue
            if params and not params.startswith("(") and "only available" not in params.lower() and "info requests" not in params.lower():
                example = f"{entity_name}({params})"
            else:
                example = entity_name if not params else f"{entity_name}({params})"
            rows.append(
                {
                    "entity_name": entity_name,
                    "allowed_parameters": params,
                    "example_usage": example,
                    "league": league,
                }
            )

    rows = sorted_unique_dict_rows(rows, ("league", "entity_name"))
    return rows, notes, used


def extract_attributes(workbooks: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    rows: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []

    mapping = [
        ("Player-Coach Attributes", "player/coach"),
        ("Team Attributes", "team"),
        ("Time Attribute", "time"),
    ]

    for league in ("MLB", "NHL"):
        for sheet_name, fallback_entity in mapping:
            sheet = find_sheet(workbooks[league], sheet_name)
            if sheet is None:
                notes.append(f"Missing expected sheet `{sheet_name}` in `{workbooks[league].file_name}`.")
                continue
            used.append(sheet)

            for row in sheet.rows:
                keys = row_keys(row)
                attribute_name = clean(
                    keys.get("player_coach_attribute")
                    or keys.get("player_coach_measure")
                    or keys.get("team_attribute")
                    or keys.get("team_measure")
                    or keys.get("formatter")
                    or ""
                )
                if not attribute_name:
                    continue
                rows.append(
                    {
                        "attribute_name": attribute_name,
                        "entity_type": clean(keys.get("entity_supported") or fallback_entity),
                        "league": league,
                        "notes": clean(keys.get("main") or keys.get("function") or ""),
                    }
                )

    rows = sorted_unique_dict_rows(rows, ("league", "entity_type", "attribute_name"))
    return rows, notes, used


def extract_formatters(workbooks: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    rows: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []

    for league in ("MLB", "NHL"):
        sheet = find_sheet(workbooks[league], "Formatters")
        if sheet is None:
            notes.append(f"Missing expected sheet `Formatters` in `{workbooks[league].file_name}`.")
            continue
        used.append(sheet)

        for row in sheet.rows:
            keys = row_keys(row)
            formatter_name = clean(keys.get("formatter", ""))
            if not formatter_name:
                continue
            if formatter_name.startswith("|"):
                usage_context = "pipe_formatter"
            elif formatter_name.startswith("(") and formatter_name.endswith(")"):
                usage_context = "measure_scope_modifier"
            else:
                usage_context = "formatter"
            rows.append(
                {
                    "formatter_name": formatter_name,
                    "description": clean(keys.get("function", "")),
                    "usage_context": usage_context,
                    "league": league,
                }
            )

    rows = sorted_unique_dict_rows(rows, ("league", "formatter_name"))
    return rows, notes, used


def extract_query_skeletons(workbooks: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    rows: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []

    for league in ("MLB", "NHL"):
        sheet = find_sheet(workbooks[league], "Available Queries")
        if sheet is None:
            notes.append(f"Missing expected sheet `Available Queries` in `{workbooks[league].file_name}`.")
            continue
        used.append(sheet)

        for row in sheet.rows:
            keys = row_keys(row)
            qtype = clean(keys.get("type", "")).lower()
            query = clean(keys.get("query", ""))
            info = clean(keys.get("info", ""))
            if not query.startswith("{{"):
                continue
            rows.append(
                {
                    "query_family": qtype,
                    "skeleton": query,
                    "example": info,
                    "league": league,
                }
            )

    rows = sorted_unique_dict_rows(rows, ("league", "query_family", "skeleton"))
    return rows, notes, used

def sheet_purpose(sheet_name: str) -> str:
    name = sheet_name.lower()
    if name == "entities":
        return "Entity tokens and parameter signatures"
    if name == "category_to_measure":
        return "General measure syntax catalog"
    if name == "category_to_measure_pitcher":
        return "Pitcher-specific measure syntax catalog (MLB)"
    if name == "category_to_measure_goalie":
        return "Goalie-specific measure syntax catalog (NHL)"
    if name == "qualifier_to_filter":
        return "Filter/qualifier syntax and parameter forms"
    if name == "player-coach attributes":
        return "Player and coach attribute names"
    if name == "team attributes":
        return "Team attribute names"
    if name == "time attribute":
        return "Time attribute and relative selectors"
    if name == "formatters":
        return "Formatter names and descriptions"
    if name == "additional measures":
        return "Extra standalone measure tokens"
    if name == "available queries":
        return "Query-family skeletons and examples"
    return "Unmapped sheet"


def write_inventory(workbooks: Dict[str, Workbook]) -> None:
    lines: List[str] = []
    lines.append("# OnAir Workbook Inventory")
    lines.append("")
    lines.append("## Purpose / Overview")
    lines.append("")
    lines.append("Deterministic inventory of workbook sheet order, headers, and sheet-level purpose for the v3 reference sources.")
    lines.append("")

    for league in ("MLB", "NHL"):
        wb = workbooks[league]
        lines.append(f"## {league} - `{wb.file_name}`")
        lines.append("")
        lines.append("| order | sheet_name | header_row | headers | data_row_count | purpose |")
        lines.append("| --- | --- | --- | --- | --- | --- |")
        for sheet in wb.sheets:
            headers = ", ".join(sheet.headers) if sheet.headers else "(none)"
            lines.append(
                f"| {sheet.order} | {escape_md(sheet.name)} | {sheet.header_row_index} | {escape_md(headers)} | {len(sheet.rows)} | {escape_md(sheet_purpose(sheet.name))} |"
            )
        lines.append("")

    lines.append("## MLB vs NHL structural differences")
    lines.append("")
    lines.append("- MLB `Entities` header is `Entity Syntax`; NHL uses `Entity`.")
    lines.append("- MLB measure headers use `Measure Syntax`; NHL uses `Measure`.")
    lines.append("- MLB filter header is `Filter (Qualifier) Syntax`; NHL uses `Qualifier/Filter`.")
    lines.append("- Player/coach attribute header differs: MLB `Player/Coach Attribute`, NHL `Player/Coach Measure`.")
    lines.append("- Team attribute header differs: MLB `Team Attribute`, NHL `Team Measure`.")
    lines.append("- MLB has `CATEGORY_TO_MEASURE_PITCHER`; NHL has `CATEGORY_TO_MEASURE_GOALIE`.")
    lines.append("")

    (DOCS_DIR / "workbook-inventory.md").write_text("\n".join(lines), encoding="utf-8")


def write_readme() -> None:
    content = """# OnAir Semantic Reference Layer

## Purpose / Overview

This directory contains read-only semantic reference artifacts extracted from the OnAir v3 workbook references.

## Source material location

- `docs/onair/source/MLB_OnAir_v3_Stat_Syntax_v3.xlsx`
- `docs/onair/source/NHL_OnAir_v3_Stat_Syntax_v3.xlsx`

## Why this documentation exists

The extracted references provide deterministic source-grounded inputs for:
- semantic dictionaries
- qualifier alias normalization
- candidate-resolution explainability
- future SmartStat plan grammar and Plan Engine preparation

## Included artifacts

- `workbook-inventory.md`
- `measure-dictionary.md`
- `filter-grammar-dictionary.md`
- `alias-map.qualifiers.md`
- `entity-dictionary.md`
- `attribute-dictionary.md`
- `formatter-dictionary.md`
- `query-skeletons.md`
- `onair_semantic_grammar.md`

## Explicit non-scope

- No SmartStat runtime behavior changes
- No resolver/planner execution changes
- No runtime integration dependencies

This layer is documentation and extraction tooling only.
"""
    (DOCS_DIR / "README.md").write_text(content, encoding="utf-8")


def write_doc(
    *,
    path: Path,
    title: str,
    overview: str,
    coverage: Sequence[str],
    notes: Sequence[str],
    headers: Sequence[str],
    rows: Sequence[Dict[str, str]],
    relevance: Sequence[str],
) -> None:
    parts: List[str] = []
    parts.append(f"# {title}")
    parts.append("")
    parts.append("## Purpose / Overview")
    parts.append("")
    parts.append(overview)
    parts.append("")
    parts.append("## Source workbook coverage")
    parts.append("")
    for item in coverage:
        parts.append(f"- {item}")
    parts.append("")
    parts.append("## Extraction notes / normalization notes")
    parts.append("")
    if notes:
        for note in notes:
            parts.append(f"- {note}")
    else:
        parts.append("- No additional normalization beyond deterministic header alignment and whitespace cleanup.")
    parts.append("")
    parts.append("## Extracted reference")
    parts.append("")
    if rows:
        parts.append(markdown_table(headers, rows))
    else:
        parts.append("_No rows extracted from mapped sheets._")
    parts.append("")
    parts.append("## SmartStat Relevance")
    parts.append("")
    for item in relevance:
        parts.append(f"- {item}")
    parts.append("")
    path.write_text("\n".join(parts), encoding="utf-8")


def write_semantic_grammar(
    *,
    workbooks: Dict[str, Workbook],
    queries: Sequence[Dict[str, str]],
    entities: Sequence[Dict[str, str]],
    filters: Sequence[Dict[str, str]],
    measures: Sequence[Dict[str, str]],
    attributes: Sequence[Dict[str, str]],
    formatters: Sequence[Dict[str, str]],
) -> None:
    query_families = sorted({q["query_family"] for q in queries})

    core_examples = [
        q for q in queries if q["query_family"] in {"info", "stats"}
    ]
    core_examples = core_examples[:8]

    entity_params = sorted({e["allowed_parameters"] for e in entities if clean(e["allowed_parameters"])})
    filter_param_examples = sorted({f["parameter_examples"] for f in filters if clean(f["parameter_examples"])})

    lines: List[str] = []
    lines.append("# OnAir Semantic Grammar (Canonical Phase 0 Reference)")
    lines.append("")
    lines.append("## Purpose / Overview")
    lines.append("")
    lines.append(
        "Canonical semantic grammar reference synthesized from the MLB/NHL OnAir v3 workbook sheets. "
        "This document defines stable query/component composition patterns for SmartStat semantic tooling."
    )
    lines.append("")
    lines.append("## Source workbook coverage")
    lines.append("")
    lines.append(f"- MLB workbook: `{workbooks['MLB'].file_name}`")
    lines.append(f"- NHL workbook: `{workbooks['NHL'].file_name}`")
    lines.append("- Source sheets used: `Entities`, `CATEGORY_TO_MEASURE*`, `QUALIFIER_TO_FILTER`, `Player-Coach Attributes`, `Team Attributes`, `Time Attribute`, `Formatters`, `Available Queries`")
    lines.append("")
    lines.append("## Extraction notes / normalization notes")
    lines.append("")
    lines.append("- Grammar statements are grounded in explicit workbook syntax/examples.")
    lines.append("- Core SmartStat planning focus remains `info` and `stats` query families; other observed families are listed as workbook-observed extensions.")
    lines.append("- No runtime resolution/execution semantics are inferred in this phase.")
    lines.append("")

    lines.append("## Query families")
    lines.append("")
    lines.append("### Core families")
    lines.append("")
    lines.append("- `info`")
    lines.append("- `stats`")
    lines.append("")
    lines.append("### Additional workbook-observed families")
    lines.append("")
    for fam in query_families:
        if fam not in {"info", "stats"}:
            lines.append(f"- `{fam}`")
    lines.append("")

    lines.append("## Semantic component classes")
    lines.append("")
    lines.append("- `entity`: actor/scope token from `Entities` (examples: `player`, `team`, `coach`, `time`).")
    lines.append("- `attribute`: `info`-oriented field tokens from attribute sheets.")
    lines.append("- `measure`: stats-valued token from `CATEGORY_TO_MEASURE*` and `Additional Measures`.")
    lines.append("- `filter`: qualifier token from `QUALIFIER_TO_FILTER`.")
    lines.append("- `formatter`: post-expression transform token from `Formatters`.")
    lines.append("")

    lines.append("## Canonical composition patterns")
    lines.append("")
    lines.append("- `{{info.entity.attribute}}`")
    lines.append("- `{{stats.entity.filter.measure}}`")
    lines.append("- Formatter application observed in workbook examples: `{{ ... | formatter }}`")
    lines.append("")
    lines.append("Representative workbook examples:")
    lines.append("")
    for item in core_examples:
        lines.append(f"- `{item['skeleton']}` ({item['league']})")
    lines.append("")

    lines.append("## Argument grammar (source-grounded)")
    lines.append("")
    lines.append("| argument_shape | description | source evidence |")
    lines.append("| --- | --- | --- |")
    lines.append("| `TRICODE` | Team shorthand code parameter | Entities `team(TRICODE)`; filters like `on_team(TRICODE)` |")
    lines.append("| `##` | Integer slot (example: player index/id parameter) | Entities `player(TRICODE, ##)` |")
    lines.append("| `##:##` | Clock-time parameter | Filter examples for `game_clock` |")
    lines.append("| `#-#` | Range interval | Filters like `innings(7-9)` examples |")
    lines.append("| `<#`, `>#`, `=#`, `<=#`, `=>#` | Comparison operators | Margin/comparison filter examples |")
    lines.append("| `last#`, `prev#` | Relative time/count selectors | `season`, `month`, `postseason` available parameters |")
    lines.append("| enum lists | Explicit token sets | Available-parameter lists in `QUALIFIER_TO_FILTER` |")
    lines.append("")

    lines.append("## League / subtype deltas")
    lines.append("")
    lines.append("- MLB adds pitcher-specific measure sheet: `CATEGORY_TO_MEASURE_PITCHER`.")
    lines.append("- NHL adds goalie-specific measure sheet: `CATEGORY_TO_MEASURE_GOALIE`.")
    lines.append("- MLB entities include `batter` and `pitcher`; NHL entity sheet omits those tokens.")
    lines.append("- Header naming differs across leagues (`Measure Syntax` vs `Measure`, `Filter (Qualifier) Syntax` vs `Qualifier/Filter`) and is normalized in extraction output.")
    lines.append("")

    lines.append("## SmartStat Relevance")
    lines.append("")
    lines.append("- Provides canonical semantic dictionaries for entities/measures/filters/attributes/formatters.")
    lines.append("- Provides source-grounded qualifier alias normalization scaffolding.")
    lines.append("- Provides deterministic input structure for candidate-resolution explainability.")
    lines.append("- Provides grammar scaffolding for later planner and Plan Engine phases.")
    lines.append("- Keeps runtime behavior unchanged while formalizing the semantic language surface.")
    lines.append("")

    lines.append("## Source evidence snapshots")
    lines.append("")
    lines.append(f"- Entity parameter forms observed: {', '.join(entity_params[:12]) if entity_params else '(none)'}")
    lines.append(f"- Filter parameter evidence examples (sample): {', '.join(filter_param_examples[:8]) if filter_param_examples else '(none)'}")
    lines.append(f"- Measure rows extracted: {len(measures)}")
    lines.append(f"- Attribute rows extracted: {len(attributes)}")
    lines.append(f"- Formatter rows extracted: {len(formatters)}")
    lines.append("")

    (DOCS_DIR / "onair_semantic_grammar.md").write_text("\n".join(lines), encoding="utf-8")

def main() -> None:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)

    missing = [str(path) for path in SOURCES.values() if not path.exists()]
    if missing:
        raise SystemExit(f"Missing required source workbook(s): {missing}")

    workbooks = {league: load_workbook(path, league) for league, path in SOURCES.items()}

    measures, measure_notes, _measure_sheets = extract_measures(workbooks)
    filters, filter_notes, _filter_sheets = extract_filters(workbooks)
    aliases, alias_notes, _alias_sheets = extract_alias_map(workbooks)
    entities, entity_notes, _entity_sheets = extract_entities(workbooks)
    attributes, attribute_notes, _attribute_sheets = extract_attributes(workbooks)
    formatters, formatter_notes, _formatter_sheets = extract_formatters(workbooks)
    query_skeletons, query_notes, _query_sheets = extract_query_skeletons(workbooks)

    write_readme()
    write_inventory(workbooks)

    write_doc(
        path=DOCS_DIR / "measure-dictionary.md",
        title="OnAir Measure Dictionary",
        overview="Deterministic measure reference extracted from category-to-measure and additional-measure source sheets.",
        coverage=[
            f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}",
            "Sheets used: CATEGORY_TO_MEASURE, CATEGORY_TO_MEASURE_PITCHER, CATEGORY_TO_MEASURE_GOALIE, Additional Measures",
        ],
        notes=[
            "Header normalization: `Measure Syntax` and `Measure` are normalized to `measure_name`.",
            "Subtype is assigned from source sheet role: `general`, `pitcher`, `goalie`, `additional`.",
            *measure_notes,
        ],
        headers=["measure_name", "description", "league", "subtype", "source_sheet"],
        rows=measures,
        relevance=[
            "Provides stable measure vocabulary for semantic dictionaries.",
            "Improves deterministic candidate-resolution context for stat terms.",
            "Supplies planner grammar measure slots without runtime coupling.",
        ],
    )

    write_doc(
        path=DOCS_DIR / "filter-grammar-dictionary.md",
        title="OnAir Filter Grammar Dictionary",
        overview="Deterministic filter/qualifier grammar reference from QUALIFIER_TO_FILTER sheets.",
        coverage=[
            f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}",
            "Sheets used: QUALIFIER_TO_FILTER",
        ],
        notes=[
            "Header normalization: `Filter (Qualifier) Syntax` and `Qualifier/Filter` are normalized to `filter_name`.",
            "`parameter_type` is conservatively classified from explicit available-parameter text and filter naming.",
            *filter_notes,
        ],
        headers=["filter_name", "parameter_type", "parameter_examples", "league", "notes"],
        rows=filters,
        relevance=[
            "Defines deterministic filter grammar tokens and parameter shapes.",
            "Supports explainable qualifier parsing and normalization.",
            "Provides source-grounded planner grammar filter slots.",
        ],
    )

    write_doc(
        path=DOCS_DIR / "alias-map.qualifiers.md",
        title="OnAir Qualifier Phrase Alias Map",
        overview="Conservative alias phrase mapping to canonical qualifier/filter forms derived from explicit syntax/parameter evidence.",
        coverage=[
            f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}",
            "Sheets used: QUALIFIER_TO_FILTER",
        ],
        notes=[
            "v3 sheets do not provide a dedicated alias column; mappings here are marked as inferred when derived from explicit parameter semantics.",
            "No free-form alias invention is performed beyond strongly indicated parameterized forms.",
            *alias_notes,
        ],
        headers=["alias_phrase", "canonical_filter", "canonical_parameter", "league", "example", "notes"],
        rows=aliases,
        relevance=[
            "Seeds deterministic qualifier alias normalization.",
            "Improves candidate-resolution explainability for human phrase variants.",
            "Provides a controlled synonym layer for future semantic/planner work.",
        ],
    )

    write_doc(
        path=DOCS_DIR / "entity-dictionary.md",
        title="OnAir Entity Dictionary",
        overview="Entity token reference with allowed parameter forms extracted from Entities sheets.",
        coverage=[
            f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}",
            "Sheets used: Entities",
        ],
        notes=[
            "Header normalization: `Entity Syntax` and `Entity` are normalized to `entity_name`.",
            "Example usage is rendered deterministically from entity name and source parameter text.",
            *entity_notes,
        ],
        headers=["entity_name", "allowed_parameters", "example_usage", "league"],
        rows=entities,
        relevance=[
            "Defines entity tokens for semantic dictionaries.",
            "Supports deterministic semantic selection and context scoping.",
            "Feeds planner grammar entity positions without runtime behavior changes.",
        ],
    )

    write_doc(
        path=DOCS_DIR / "attribute-dictionary.md",
        title="OnAir Attribute Dictionary",
        overview="Attribute token reference across player/coach, team, and time attribute source sheets.",
        coverage=[
            f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}",
            "Sheets used: Player-Coach Attributes, Team Attributes, Time Attribute",
        ],
        notes=[
            "Header normalization aligns `Player/Coach Attribute` vs `Player/Coach Measure` and `Team Attribute` vs `Team Measure`.",
            "`Time Attribute` rows are included under `entity_type=time`.",
            *attribute_notes,
        ],
        headers=["attribute_name", "entity_type", "league", "notes"],
        rows=attributes,
        relevance=[
            "Defines attribute vocabulary for semantic dictionaries.",
            "Supports explainable entity->attribute resolution paths.",
            "Supplies planner grammar attribute slots with source provenance.",
        ],
    )

    write_doc(
        path=DOCS_DIR / "formatter-dictionary.md",
        title="OnAir Formatter Dictionary",
        overview="Formatter token reference extracted from Formatters sheets.",
        coverage=[
            f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}",
            "Sheets used: Formatters",
        ],
        notes=[
            "`usage_context` classification: `pipe_formatter`, `measure_scope_modifier`, or `formatter`.",
            *formatter_notes,
        ],
        headers=["formatter_name", "description", "usage_context", "league"],
        rows=formatters,
        relevance=[
            "Defines formatter vocabulary and usage shape.",
            "Supports explainable expression post-processing semantics.",
            "Provides planner grammar formatter references without runtime coupling.",
        ],
    )

    write_doc(
        path=DOCS_DIR / "query-skeletons.md",
        title="OnAir Query Skeleton Grammar",
        overview="Workbook-grounded query template reference extracted from Available Queries sheets.",
        coverage=[
            f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}",
            "Sheets used: Available Queries",
        ],
        notes=[
            "Rows are included only when `Query` contains an explicit `{{...}}` template.",
            "Query families are taken directly from the workbook `Type` column (normalized to lowercase).",
            *query_notes,
        ],
        headers=["query_family", "skeleton", "example", "league"],
        rows=query_skeletons,
        relevance=[
            "Provides canonical query-shape references for semantic parsing.",
            "Supports candidate-resolution explainability by query-family context.",
            "Supplies baseline planner grammar templates for future phases.",
        ],
    )

    write_semantic_grammar(
        workbooks=workbooks,
        queries=query_skeletons,
        entities=entities,
        filters=filters,
        measures=measures,
        attributes=attributes,
        formatters=formatters,
    )

    print("OnAir v3 extraction complete")
    print(f"Docs written to: {DOCS_DIR}")
    print(f"Measure rows: {len(measures)}")
    print(f"Filter rows: {len(filters)}")
    print(f"Alias rows: {len(aliases)}")
    print(f"Entity rows: {len(entities)}")
    print(f"Attribute rows: {len(attributes)}")
    print(f"Formatter rows: {len(formatters)}")
    print(f"Query rows: {len(query_skeletons)}")


if __name__ == "__main__":
    main()
