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
