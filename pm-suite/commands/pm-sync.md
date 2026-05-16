---
description: Sync project artifacts to/from an external system via an adapter
argument-hint: notion|github <project>
---

Run a storage-adapter sync for `pm-suite/projects/$2/` using adapter `$1`.

Steps:

1. Confirm `$1` is one of: `notion`, `github`, `local-markdown` (no-op).
2. Read `pm-suite/adapters/$1.md` for the adapter's operations and the MCP tools it uses.
3. Confirm the relevant MCP server is connected. Use `ToolSearch` with the query `$1` to list available tools. If no relevant tools are found, tell the user and stop.
4. Read the link map file (`.adapter-$1.json` in the project directory) if it exists; otherwise the sync will create one.
5. Ask the user which direction:
   - **push** (Suite → External): mirror local artifacts to the external system.
   - **pull** (External → Suite): pull updates into local markdown (surface diffs; do not auto-overwrite).
   - **diff** (compare only).

6. Execute the chosen direction artifact-by-artifact. For each artifact:
   - Report the action taken (created / updated / skipped / conflict).
   - For conflicts: save the diff to `.adapter-$1-diff.md` and tell the user to resolve via `/pm-decide`.

7. After the sync, update the link map file.

## Conventions

- **Source of truth is markdown** for plans, decisions, audits, principles, charter, business case.
- **Source of truth is the external system** for execution state (e.g., GitHub Issue status) when the team works there day-to-day.
- Conflicts are not auto-resolved.
- Adapters must never silently change the external schema.

Cite Principle 1 (Stewardship), 5 (Systems thinking), 7 (Tailoring).
