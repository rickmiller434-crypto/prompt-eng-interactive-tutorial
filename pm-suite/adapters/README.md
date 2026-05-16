# Storage Adapters

The PM Suite is **storage-agnostic**. Agents read from and write to `pm-suite/projects/<project>/` as plain markdown by default. Adapters are thin layers that mirror those artifacts to (or from) other systems of record — without changing the agents.

## Adapter contract

Every adapter implements four operations:

| Operation | Direction | Purpose |
|-----------|-----------|---------|
| `push` | Suite → System | Send local artifact to the external system |
| `pull` | System → Suite | Pull updates from the external system into local markdown |
| `link` | bidirectional | Maintain a stable mapping (file ↔ external ID) |
| `diff` | comparison | Show divergence between local and external |

Adapters do not own the artifacts — the markdown files in `pm-suite/projects/<project>/` remain the source of truth. The external system is a projection.

## Available adapters

| Adapter | Status | Use when… | File |
|---------|--------|------------|------|
| **local-markdown** | ✅ Default | Always — this is the source of truth | `local-markdown.md` |
| **notion** | ✅ Spec'd | Stakeholders / sponsors live in Notion | `notion.md` |
| **github-projects** | ✅ Spec'd | Team works primarily in GitHub Issues/Projects | `github-projects.md` |

## Choosing an adapter

The Tailoring Record (`tailoring.md` → Section H Q36 — Mandatory tools) drives the choice. Multiple adapters can run in parallel: e.g., backlog mirrored to GitHub Projects for the engineering team while charter / status / decisions mirror to Notion for leadership.

## Conflict policy

When local and external diverge:

1. Markdown is canonical for **plans, decisions, audits, principles**.
2. External is canonical for **execution state of items the team works on there** (e.g., issue status when the team uses GitHub Issues live).
3. The `diff` operation surfaces the conflict; resolution is a `pm-work` decision logged in `decisions/decision-log.md`.

## Adapter implementations

These adapter specs describe **what** each adapter does, **which MCP server tools** to use, and **how** to invoke them from Claude Code. They are not standalone scripts — adapters run inside the Claude Code session via the relevant MCP servers (Notion MCP, GitHub MCP) and the file tools.
