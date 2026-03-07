#!/usr/bin/env python3
"""
Deterministic Phase 0 OnAir semantic reference extraction.
Standard library only: zipfile + XML workbook parsing.
"""

from __future__ import annotations

import json
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
OUT_JSON = DOCS_DIR / "_extracted" / "onair_reference.extracted.json"

SOURCES = {
    "MLB": SOURCE_DIR / "MLB_OnAir_v3_Stat_Syntax_v2.xlsx",
    "NHL": SOURCE_DIR / "NHL_OnAir_v3_Stat_Syntax_v2.xlsx",
}


def clean(value: str) -> str:
    return value.strip().replace("\n", " ").replace("\r", " ")


def header_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", clean(value).lower()).strip("_")


def escape_md(value: str) -> str:
    return value.replace("|", "\\|")


def join_nonempty(values: Iterable[str], sep: str = "; ") -> str:
    items = [clean(v) for v in values if clean(v)]
    return sep.join(items)


def parse_filter_and_param(value: str) -> Tuple[str, str]:
    txt = clean(value)
    m = re.match(r"^([^(]+)\(([^)]+)\)$", txt)
    if not m:
        return txt, ""
    return clean(m.group(1)), clean(m.group(2))


def classify_parameter_type(value: str) -> str:
    txt = clean(value).lower()
    if not txt:
        return "none"
    if "tricode" in txt:
        return "team_tricode"
    if "year" in txt:
        return "year"
    if "margin" in txt:
        return "score_margin"
    if "clock" in txt or re.search(r"\b\d{1,2}:\d{2}\b", txt):
        return "clock_time"
    if any(t in txt for t in ["player", "team", "entity"]):
        return "entity_reference"
    if any(t in txt for t in ["<", ">", "less than", "greater than", "over", "under"]):
        return "comparison"
    if any(t in txt for t in [" to ", " through ", " - "]):
        return "range"
    if "," in txt or " or " in txt or any(t in txt for t in ["lead", "trail", "tie", "before", "after"]):
        return "enum"
    return "unknown"


def markdown_table(headers: Sequence[str], rows: Sequence[Dict[str, str]]) -> str:
    out = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    for row in rows:
        out.append("| " + " | ".join(escape_md(str(row.get(h, ""))) for h in headers) + " |")
    return "\n".join(out)


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


def read_rows(sheet_xml: ET.Element, shared: List[str]) -> List[List[str]]:
    rows: List[List[str]] = []
    for row in sheet_xml.findall("m:sheetData/m:row", NS):
        cells: List[str] = []
        for c in row.findall("m:c", NS):
            ctype = c.attrib.get("t")
            v = c.find("m:v", NS)
            if ctype == "inlineStr":
                cells.append(clean("".join(t.text or "" for t in c.findall(".//m:t", NS))))
                continue
            if v is None:
                cells.append("")
                continue
            raw = v.text or ""
            if ctype == "s":
                try:
                    cells.append(clean(shared[int(raw)]))
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
    hdr = headers + [""] * (width - len(headers))
    out: List[Dict[str, str]] = []
    for row in rows[header_idx + 1 :]:
        vals = row + [""] * (width - len(row))
        if not any(clean(v) for v in vals):
            continue
        rec: Dict[str, str] = {}
        for i, value in enumerate(vals):
            key = hdr[i] if hdr[i] else f"column_{i+1}"
            rec[key] = clean(value)
        out.append(rec)
    return out


def load_workbook(path: Path, league: str) -> Workbook:
    with zipfile.ZipFile(path) as zf:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        rel_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rel_xml.findall("pr:Relationship", NS)
        }

        shared: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst.findall("m:si", NS):
                shared.append(clean("".join(t.text or "" for t in si.findall(".//m:t", NS))))

        sheets: List[Sheet] = []
        for order, sheet_node in enumerate(wb_xml.findall("m:sheets/m:sheet", NS), start=1):
            name = sheet_node.attrib["name"]
            rid = sheet_node.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = rel_map[rid]
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            sheet_xml = ET.fromstring(zf.read(target))
            rows = read_rows(sheet_xml, shared)
            header_idx, headers = detect_header(rows)
            sheets.append(
                Sheet(
                    name=name,
                    order=order,
                    path=target,
                    headers=headers,
                    rows=build_records(rows, header_idx, headers),
                    header_row_index=header_idx + 1,
                )
            )
    return Workbook(league=league, file_name=path.name, sheets=sheets)


def find_sheet(wb: Workbook, sheet_name: str) -> Sheet | None:
    target = sheet_name.lower()
    for sheet in wb.sheets:
        if sheet.name.lower() == target:
            return sheet
    return None


def row_keys(row: Dict[str, str]) -> Dict[str, str]:
    return {header_key(k): v for k, v in row.items()}


def split_aliases(value: str) -> List[str]:
    return [clean(part) for part in value.split(",") if clean(part)]


def used_sheet_labels(sheets: Sequence[Sheet]) -> str:
    return ", ".join(f"{s.name} ({s.order})" for s in sheets) if sheets else "(none)"


def extract_measures(wbs: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    data: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []
    mapping = [
        ("MLB", "CATEGORY_TO_MEASURE", "general"),
        ("MLB", "CATEGORY_TO_MEASURE_PITCHER", "pitcher"),
        ("NHL", "CATEGORY_TO_MEASURE", "general"),
        ("NHL", "CATEGORY_TO_MEASURE_GOALIE", "goalie"),
        ("MLB", "Additional Measures", "additional"),
        ("NHL", "Additional Measures", "additional"),
    ]
    for league, sheet_name, subtype in mapping:
        sheet = find_sheet(wbs[league], sheet_name)
        if sheet is None:
            notes.append(f"Missing expected sheet `{sheet_name}` in `{wbs[league].file_name}`.")
            continue
        used.append(sheet)
        for row in sheet.rows:
            keys = row_keys(row)
            measure_name = clean(keys.get("measure") or keys.get("value") or keys.get("syntax") or "")
            if not measure_name:
                continue
            data.append(
                {
                    "measure_name": measure_name,
                    "description": clean(keys.get("description") or ""),
                    "league": league,
                    "subtype": subtype,
                    "source_sheet": sheet.name,
                    "notes": "",
                }
            )
    data.sort(key=lambda r: (r["league"], r["subtype"], r["measure_name"]))
    return data, notes, used


def extract_filters(wbs: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    data: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []
    for league in ("MLB", "NHL"):
        sheet = find_sheet(wbs[league], "QUALIFIER_TO_FILTER")
        if sheet is None:
            notes.append(f"Missing `QUALIFIER_TO_FILTER` in `{wbs[league].file_name}`.")
            continue
        used.append(sheet)
        for row in sheet.rows:
            keys = row_keys(row)
            filter_name = clean(keys.get("qualifier_filter") or keys.get("filter") or "")
            if not filter_name:
                continue
            available = clean(keys.get("available_parameters") or "")
            examples = join_nonempty(
                [keys.get("example_a", ""), keys.get("example_b", ""), keys.get("example_c", ""), keys.get("example_d", "")]
            )
            data.append(
                {
                    "filter_name": filter_name,
                    "parameter_type": classify_parameter_type(available),
                    "parameter_examples": join_nonempty([available, examples]),
                    "league": league,
                    "source_sheet": sheet.name,
                    "notes": "Alias column present." if clean(keys.get("aliases", "")) else "",
                }
            )
    data.sort(key=lambda r: (r["league"], r["filter_name"]))
    return data, notes, used


def extract_aliases(wbs: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    data: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []
    for league in ("MLB", "NHL"):
        sheet = find_sheet(wbs[league], "QUALIFIER_TO_FILTER")
        if sheet is None:
            continue
        used.append(sheet)
        for row in sheet.rows:
            keys = row_keys(row)
            aliases = split_aliases(keys.get("aliases", ""))
            canonical = clean(keys.get("qualifier_filter") or keys.get("filter") or "")
            if not aliases or not canonical:
                continue
            canon_filter, canon_param = parse_filter_and_param(canonical)
            example = join_nonempty(
                [keys.get("example_a", ""), keys.get("example_b", ""), keys.get("example_c", ""), keys.get("example_d", "")]
            )
            for alias in aliases:
                data.append(
                    {
                        "alias_phrase": alias,
                        "canonical_filter": canon_filter,
                        "canonical_parameter": canon_param,
                        "league": league,
                        "source_sheet": sheet.name,
                        "example": example,
                        "notes": "Grounded in `Aliases` column.",
                    }
                )
    if not data:
        notes.append("No alias phrases were present in source sheets.")
    data.sort(key=lambda r: (r["league"], r["canonical_filter"], r["alias_phrase"]))
    return data, notes, used


def extract_entities(wbs: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    data: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []
    for league in ("MLB", "NHL"):
        sheet = find_sheet(wbs[league], "Entities")
        if sheet is None:
            notes.append(f"Missing `Entities` in `{wbs[league].file_name}`.")
            continue
        used.append(sheet)
        for row in sheet.rows:
            keys = row_keys(row)
            entity_name = clean(keys.get("entity", ""))
            if not entity_name:
                continue
            data.append(
                {
                    "entity_name": entity_name,
                    "allowed_parameters": clean(keys.get("available_parameters", "")),
                    "example_usage": join_nonempty([keys.get("example_a", ""), keys.get("example_b", ""), keys.get("example_c", ""), keys.get("example_d", "")]),
                    "league": league,
                    "source_sheet": sheet.name,
                    "notes": join_nonempty([keys.get("key", "")]),
                }
            )
    data.sort(key=lambda r: (r["league"], r["entity_name"]))
    return data, notes, used


def extract_attributes(wbs: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    data: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []
    mapping = [
        ("Player-Coach Attributes", "player_coach"),
        ("Team Attributes", "team"),
        ("Time Attribute", "time"),
    ]
    for league in ("MLB", "NHL"):
        for sheet_name, entity_fallback in mapping:
            sheet = find_sheet(wbs[league], sheet_name)
            if sheet is None:
                notes.append(f"Missing `{sheet_name}` in `{wbs[league].file_name}`.")
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
                data.append(
                    {
                        "attribute_name": attribute_name,
                        "entity_type": clean(keys.get("entity_supported") or entity_fallback),
                        "league": league,
                        "source_sheet": sheet.name,
                        "notes": clean(keys.get("function", "")),
                    }
                )
    data.sort(key=lambda r: (r["league"], r["entity_type"], r["attribute_name"]))
    return data, notes, used


def extract_formatters(wbs: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    data: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []
    for league in ("MLB", "NHL"):
        sheet = find_sheet(wbs[league], "Formatters")
        if sheet is None:
            notes.append(f"Missing `Formatters` in `{wbs[league].file_name}`.")
            continue
        used.append(sheet)
        for row in sheet.rows:
            keys = row_keys(row)
            formatter = clean(keys.get("formatter", ""))
            if not formatter:
                continue
            data.append(
                {
                    "formatter_name": formatter,
                    "description": clean(keys.get("function") or keys.get("description") or ""),
                    "usage_context": "pipe_formatter" if formatter.startswith("|") else "measure_modifier",
                    "league": league,
                    "source_sheet": sheet.name,
                    "notes": "",
                }
            )
    data.sort(key=lambda r: (r["league"], r["formatter_name"]))
    return data, notes, used


def extract_queries(wbs: Dict[str, Workbook]) -> Tuple[List[Dict[str, str]], List[str], List[Sheet]]:
    data: List[Dict[str, str]] = []
    notes: List[str] = []
    used: List[Sheet] = []
    for league in ("MLB", "NHL"):
        sheet = find_sheet(wbs[league], "Available Queries")
        if sheet is None:
            notes.append(f"Missing `Available Queries` in `{wbs[league].file_name}`.")
            continue
        used.append(sheet)
        shifted = 0
        for row in sheet.rows:
            keys = row_keys(row)
            type_col = clean(keys.get("type", ""))
            query_col = clean(keys.get("query", ""))
            info_col = clean(keys.get("info", ""))
            row_note = ""

            if not query_col and type_col.startswith("{{"):
                shifted += 1
                query_text = type_col
                query_family = "INFO" if "info." in query_text else "UNKNOWN"
                example = clean(keys.get("query", "") or info_col)
                row_note = "Query value was shifted under `Type` column in source row."
            else:
                query_text = query_col
                query_family = type_col
                example = info_col

            if not query_text.startswith("{{"):
                continue

            data.append(
                {
                    "query_family": query_family or "UNKNOWN",
                    "skeleton": query_text,
                    "example": example,
                    "league": league,
                    "source_sheet": sheet.name,
                    "notes": row_note,
                }
            )
        if shifted:
            notes.append(f"`{wbs[league].file_name}` had {shifted} shifted query row(s) in `Available Queries`.")
    data.sort(key=lambda r: (r["league"], r["query_family"], r["skeleton"]))
    return data, notes, used


def write_doc(
    path: Path,
    title: str,
    overview: str,
    coverage: Sequence[str],
    extraction_notes: Sequence[str],
    headers: Sequence[str],
    rows: Sequence[Dict[str, str]],
    relevance: Sequence[str],
    row_limit: int | None = None,
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
    for line in coverage:
        parts.append(f"- {line}")
    parts.append("")
    parts.append("## Extraction notes / normalization notes")
    parts.append("")
    if extraction_notes:
        for note in extraction_notes:
            parts.append(f"- {note}")
    else:
        parts.append("- No additional normalization beyond header alignment and whitespace cleanup.")
    parts.append("")
    parts.append("## Extracted reference")
    parts.append("")
    shown_rows = list(rows[:row_limit]) if row_limit is not None else list(rows)
    parts.append(markdown_table(headers, shown_rows) if shown_rows else "_No rows extracted from mapped sheets._")
    if row_limit is not None and len(rows) > row_limit:
        parts.append("")
        parts.append(
            f"_Table truncated to first {row_limit} rows for readability. Full extracted rows are available in `docs/onair/_extracted/onair_reference.extracted.json`._"
        )
    parts.append("")
    parts.append("## SmartStat Relevance")
    parts.append("")
    for item in relevance:
        parts.append(f"- {item}")
    parts.append("")
    path.write_text("\n".join(parts), encoding="utf-8")


def write_readme() -> None:
    content = """# OnAir Semantic Reference Layer (Phase 0)

## What This Is

This directory contains read-only semantic reference artifacts extracted from OnAir workbook references.

## Source material location

- `docs/onair/source/MLB_OnAir_v3_Stat_Syntax_v2.xlsx`
- `docs/onair/source/NHL_OnAir_v3_Stat_Syntax_v2.xlsx`

## Why these docs exist

These references support:
- semantic dictionaries
- alias normalization
- candidate-resolution explainability
- planner grammar preparation

## Explicit non-scope

- No SmartStat runtime behavior implementation
- No resolver logic implementation
- No planner execution implementation
- No runtime integration

## Artifact index

- `measure-dictionary.md`
- `filter-grammar-dictionary.md`
- `alias-map.qualifiers.md`
- `entity-dictionary.md`
- `attribute-dictionary.md`
- `formatter-dictionary.md`
- `query-skeletons.md`
- `workbook-inventory.md`
"""
    (DOCS_DIR / "README.md").write_text(content, encoding="utf-8")


def write_inventory(wbs: Dict[str, Workbook]) -> None:
    lines: List[str] = []
    lines.append("# OnAir Workbook Inventory")
    lines.append("")
    lines.append("## Purpose / Overview")
    lines.append("")
    lines.append("Deterministic inventory of actual workbook sheets, order, and header structures.")
    lines.append("")
    for league in ("MLB", "NHL"):
        wb = wbs[league]
        lines.append(f"## {league} - `{wb.file_name}`")
        lines.append("")
        lines.append("| order | sheet_name | header_row | headers | data_row_count |")
        lines.append("| --- | --- | --- | --- | --- |")
        for sheet in wb.sheets:
            headers = ", ".join(sheet.headers) if sheet.headers else "(none)"
            lines.append(
                f"| {sheet.order} | {escape_md(sheet.name)} | {sheet.header_row_index} | {escape_md(headers)} | {len(sheet.rows)} |"
            )
        lines.append("")
    lines.append("## Structural differences noted")
    lines.append("")
    lines.append("- NHL `Entities` includes `Key` and example columns; MLB `Entities` is compact.")
    lines.append("- MLB `CATEGORY_TO_MEASURE_PITCHER` uses header `SYNTAX`; other measure sheets use `VALUE`/`Measure`.")
    lines.append("- NHL `QUALIFIER_TO_FILTER` has explicit `Aliases`; MLB does not.")
    lines.append("- `Available Queries` includes shifted rows where query text appears under `Type`.")
    lines.append("- `Time Attribute` uses `Formatter`/`Function` headers and is treated as attribute reference input.")
    lines.append("")
    (DOCS_DIR / "workbook-inventory.md").write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    workbooks = {league: load_workbook(path, league) for league, path in SOURCES.items()}

    measures, measure_notes, measure_sheets = extract_measures(workbooks)
    filters, filter_notes, filter_sheets = extract_filters(workbooks)
    aliases, alias_notes, alias_sheets = extract_aliases(workbooks)
    entities, entity_notes, entity_sheets = extract_entities(workbooks)
    attributes, attribute_notes, attribute_sheets = extract_attributes(workbooks)
    formatters, formatter_notes, formatter_sheets = extract_formatters(workbooks)
    queries, query_notes, query_sheets = extract_queries(workbooks)

    json_payload = {
        "workbook_inventory": {
            league: {
                "workbook_file": wb.file_name,
                "sheets": [
                    {
                        "order": s.order,
                        "sheet_name": s.name,
                        "sheet_path": s.path,
                        "header_row": s.header_row_index,
                        "headers": s.headers,
                        "data_row_count": len(s.rows),
                    }
                    for s in wb.sheets
                ],
            }
            for league, wb in workbooks.items()
        },
        "measure_dictionary": measures,
        "filter_grammar_dictionary": filters,
        "alias_map_qualifiers": aliases,
        "entity_dictionary": entities,
        "attribute_dictionary": attributes,
        "formatter_dictionary": formatters,
        "query_skeletons": queries,
    }
    OUT_JSON.write_text(json.dumps(json_payload, indent=2, ensure_ascii=False), encoding="utf-8")

    write_readme()
    write_inventory(workbooks)

    write_doc(
        DOCS_DIR / "measure-dictionary.md",
        "OnAir Measure Dictionary",
        "Reference list of measure tokens from category/measure sheets and additional measure sheets.",
        [f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}", f"Sheets used: {used_sheet_labels(measure_sheets)}"],
        ["Header normalization: `VALUE`, `SYNTAX`, and `Measure` -> `measure_name`.", "Subtype derives from sheet role (`general`, `pitcher`, `goalie`, `additional`).", *measure_notes],
        ["measure_name", "description", "league", "subtype", "source_sheet", "notes"],
        measures,
        [
            "Supports semantic measure dictionaries.",
            "Supports candidate-resolution explainability with stable measure vocabulary.",
            "Provides planner grammar measure inputs without runtime coupling.",
        ],
        row_limit=200,
    )
    write_doc(
        DOCS_DIR / "filter-grammar-dictionary.md",
        "OnAir Filter Grammar Dictionary",
        "Reference list of filter forms and parameter signatures from QUALIFIER_TO_FILTER sheets.",
        [f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}", f"Sheets used: {used_sheet_labels(filter_sheets)}"],
        ["Header normalization: `Filter` and `Qualifier/Filter` -> `filter_name`.", "Parameter type classification is conservative and deterministic.", *filter_notes],
        ["filter_name", "parameter_type", "parameter_examples", "league", "source_sheet", "notes"],
        filters,
        [
            "Supports semantic filter dictionaries.",
            "Supports qualifier parameter explainability.",
            "Supports future planner grammar inputs without resolver/runtime implementation.",
        ],
    )
    write_doc(
        DOCS_DIR / "alias-map.qualifiers.md",
        "OnAir Qualifier Phrase Alias Map",
        "Alias phrases mapped to canonical qualifier/filter forms where alias columns are explicitly available.",
        [f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}", f"Sheets used: {used_sheet_labels(alias_sheets)}"],
        ["Only source-grounded alias rows are emitted; no free-form alias invention.", "Canonical filter/parameter parse supports structured forms like `all_star_break(after)`.", *alias_notes],
        ["alias_phrase", "canonical_filter", "canonical_parameter", "league", "source_sheet", "example", "notes"],
        aliases,
        [
            "Supports qualifier alias normalization.",
            "Improves deterministic candidate-resolution wording.",
            "Supplies planner synonym scaffolding without planner execution behavior.",
        ],
    )
    write_doc(
        DOCS_DIR / "entity-dictionary.md",
        "OnAir Entity Dictionary",
        "Entity references with supported parameters and source examples.",
        [f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}", f"Sheets used: {used_sheet_labels(entity_sheets)}"],
        ["NHL includes extra `Key`/example columns; MLB is compact.", "Example fields are merged from source `Example A..D` columns when present.", *entity_notes],
        ["entity_name", "allowed_parameters", "example_usage", "league", "source_sheet", "notes"],
        entities,
        [
            "Supports semantic entity dictionaries.",
            "Supports deterministic entity selection context for explainability.",
            "Supports planner entity grammar slots without runtime integration.",
        ],
    )
    write_doc(
        DOCS_DIR / "attribute-dictionary.md",
        "OnAir Attribute Dictionary",
        "Attribute references across player/coach, team, and time sheets.",
        [f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}", f"Sheets used: {used_sheet_labels(attribute_sheets)}"],
        ["Header normalization aligns source naming differences (`Attribute` vs `Measure`).", "`Time Attribute` source headers (`Formatter`/`Function`) are mapped conservatively into attribute rows.", *attribute_notes],
        ["attribute_name", "entity_type", "league", "source_sheet", "notes"],
        attributes,
        [
            "Supports semantic attribute dictionaries.",
            "Supports explainable entity-attribute chains.",
            "Supports planner grammar attribute slots without resolver/runtime behavior.",
        ],
    )
    write_doc(
        DOCS_DIR / "formatter-dictionary.md",
        "OnAir Formatter Dictionary",
        "Formatter references and usage context metadata from Formatters sheets.",
        [f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}", f"Sheets used: {used_sheet_labels(formatter_sheets)}"],
        ["Usage context is tagged as `pipe_formatter` for pipe-prefixed formatter names; otherwise `measure_modifier`.", *formatter_notes],
        ["formatter_name", "description", "usage_context", "league", "source_sheet", "notes"],
        formatters,
        [
            "Supports semantic formatter dictionaries.",
            "Supports explainability of expression post-processing hints.",
            "Supports planner grammar formatter references without runtime behavior.",
        ],
    )
    write_doc(
        DOCS_DIR / "query-skeletons.md",
        "OnAir Query Skeleton Grammar",
        "Canonical query-template shapes extracted from `Available Queries` sheets.",
        [f"Workbook(s): {workbooks['MLB'].file_name}, {workbooks['NHL'].file_name}", f"Sheets used: {used_sheet_labels(query_sheets)}"],
        ["Only rows containing `{{...}}` templates are included.", "Shifted source rows (query text under `Type`) are normalized and documented.", *query_notes],
        ["query_family", "skeleton", "example", "league", "source_sheet", "notes"],
        queries,
        [
            "Supports semantic query-shape dictionaries.",
            "Supports candidate-resolution context by query family.",
            "Supports planner grammar scaffolding without planner/runtime execution.",
        ],
    )

    print("OnAir extraction complete")
    print(f"Wrote docs under: {DOCS_DIR}")
    print(f"Wrote extraction JSON: {OUT_JSON}")


if __name__ == "__main__":
    main()
