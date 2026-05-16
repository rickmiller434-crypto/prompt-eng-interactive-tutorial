# Adapter — Notion

Mirrors PM Suite artifacts to Notion pages and databases. Useful when sponsors, steering committees, or non-engineering stakeholders work in Notion.

## Requirements

- Notion MCP server connected to the Claude Code session (tools prefixed `mcp__*__notion-*` — search via `ToolSearch` with query `notion`).
- A Notion workspace and team the user has access to.
- A Notion folder or top-level page to host the project workspace.

## Suggested Notion structure

```
<Project name> (page)
├── 📜 Charter (page mirroring charter.md)
├── 💰 Business case (page)
├── 🎯 Tailoring (page)
├── 🧭 Decisions (database — one row per decision-log entry)
├── 👥 Stakeholders (database — one row per stakeholder)
├── ⚠️ Risks (database — one row per risk)
├── 📋 Backlog (database — one row per backlog item)
├── 📈 Status reports (database — one row per period)
├── 🔍 Audits (database — one row per audit)
└── 🎬 Closure (page)
```

## Field mapping (representative)

| Local artifact | Notion target | Mapping |
|----------------|---------------|---------|
| `charter.md` | "Charter" page | Markdown → Notion blocks |
| `decisions/decision-log.md` rows | "Decisions" DB | Each row → Notion page with properties (Date, ID, Decision, Rationale, Level, Steward, Principles) |
| `risk-register.md` rows | "Risks" DB | Per-risk row → Notion page with properties (ID, Type, Score, Owner, Status, Review date) |
| `product-backlog.md` rows | "Backlog" DB | Each item → Notion page (Status, Priority, Size, Benefit, Stakeholder, AC) |
| `status-reports/<date>-status.md` | "Status reports" DB | Each report → Notion page |
| `audits/audit-<date>.md` | "Audits" DB | Each audit → Notion page |

## Operations

### push (Suite → Notion)

1. Locate or create the project page (search by name → `notion-search`).
2. For each artifact category, locate or create the database / page (`notion-create-database`, `notion-create-pages`).
3. Render each markdown artifact's structured rows into Notion pages with the relevant properties (`notion-create-pages` / `notion-update-page`).
4. Store the Notion page IDs locally in `pm-suite/projects/<project>/.adapter-notion.json` for `link`.

### pull (Notion → Suite)

1. For each linked Notion DB row, fetch with `notion-fetch`.
2. Convert content into the corresponding markdown table row / section.
3. Surface diffs (do not auto-overwrite — let `pm-work` resolve).

### link

Maintain `.adapter-notion.json`:

```json
{
  "project_page": "<notion_id>",
  "databases": {
    "decisions": "<id>",
    "risks": "<id>",
    "backlog": "<id>",
    "status_reports": "<id>",
    "audits": "<id>"
  },
  "row_links": {
    "decisions": { "D-001": "<page_id>", "D-002": "<page_id>" },
    "risks":     { "R1": "<page_id>", "R2": "<page_id>" }
  }
}
```

### diff

For each linked artifact, fetch the Notion version and report differences. Do not auto-merge.

## Conventions

- **Source of truth remains markdown.** Notion is a projection; conflicting edits resolve toward the markdown unless a decision-log entry says otherwise.
- **Property names match markdown column headers** wherever possible.
- **Principles footer** is included as the last block on every mirrored page.

## Invocation

Run the adapter from a Claude Code session that has the Notion MCP connected. Search MCP tools via `ToolSearch` query `notion`. Use `/pm-sync notion <project>` (Phase 6 command) to run the push/pull/diff round trip.

## Failure modes

- Notion rate limits → adapter pauses and retries (exponential backoff).
- Notion schema drift (someone changed a DB column) → adapter halts and reports; do not silently rewrite the Notion schema.
- Network failures → no partial writes; adapter is atomic per artifact.
