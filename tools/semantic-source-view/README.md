# SmartStat Semantic Source View

Phase 5 introduces a local, read-only React explorer for inspecting semantic-layer outputs.

This app:
- reads `.tools/onair_dump/index/semantic_index.json`
- renders deterministic semantic inspection views
- does not modify SmartStat runtime behavior
- does not mutate semantic index data

## Scope Covered In Phase 5

- source tree browsing (`source_tree`)
- records inspection (`entities`, `measures`, `filters`, `profiles`, `qualifiers`)
- relationship inspection (`relationships`)
- query-path inspection (`query_paths`)
- traceability inspection (`trace_index`)
- shared filtering by league, record type, source type (`ui_views` aware)

## Phase 6 Additions

- deterministic search explainability
- deterministic match ranking
- result badges for match strength/type
- record detail debug narratives
- empty-result guidance with source-oriented hints
- read-only behavior preserved

Example expectations:
- search `air_balls` prioritizes `measure:mlb:air_balls` over `measure:mlb:air_balls_percentage`
- search `playerSplits` may produce no normalized-record match and still show source/debug hints

## Phase 7 Additions

- extracted deterministic search contract helpers into a reusable module
- added lightweight deterministic ranking tests
- locked normalized-vs-source-hint search behavior for future planner/explainability work

Phase 7 checkpoint:
- commit: `30d15b6`
- tag: `semantic-view-phase7`

## Phase 8 Scaffold Start

- adds read-only semantic resolution explainability model scaffolding (`resolution-explainability.v0`)
- introduces deterministic explainability steps for selected records
- includes a minimal fixture for scaffold/demo usage
- keeps runtime apply behavior, planner execution, and SmartStat runtime integration out of scope

See:
- `RESOLUTION_EXPLAINABILITY_GUIDE.md`
- `resolution-explainability.schema.json`
- `src/data/resolutionExplainability.fixture.ts`

## Deferred

- editing semantic records
- editing query paths
- runtime SmartStat integration
- applying plans
- direct OnAir querying
- production packaging

## Local Run

From repo root:

```powershell
cd tools/semantic-source-view
npm install
npm run dev
```

Open the local Vite URL (typically `http://localhost:5173`).

The app fetches semantic data through a local dev endpoint:
- `GET /api/semantic-index`

That endpoint reads directly from:
- `.tools/onair_dump/index/semantic_index.json`

## Notes

- This tool is intended for developer exploration and future Plan Engine/source-debug workflows.
- Deterministic ordering follows the semantic index payload ordering where applicable.
