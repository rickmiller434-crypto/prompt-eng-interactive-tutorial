# Adapter — GitHub Projects + Issues

Mirrors execution-state artifacts (backlog, decisions, risks) to GitHub Issues and GitHub Projects. Useful when the engineering team works primarily in GitHub and you want the suite's artifacts to flow with their day-to-day.

## Requirements

- GitHub MCP server connected (tools prefixed `mcp__github__*`) — already available in this environment.
- A GitHub repo and a GitHub Project (Projects v2 — beta UI) created in advance.
- Repository scope authorized for the MCP server.

## What gets mirrored

| Local artifact | GitHub target | Notes |
|----------------|----------------|-------|
| `product-backlog.md` rows | One Issue per item; one Project card | Issue labels: `pm:backlog`, `pm:benefit-B1`, etc. |
| `risk-register.md` rows | One Issue per risk; label `pm:risk` | Project column reflects status |
| `decisions/decision-log.md` rows | One Issue per material decision; label `pm:decision` | Closed when superseded |
| `status-reports/<date>-status.md` | GitHub Discussion or Issue per period | Linked to milestone |
| `audits/audit-<date>.md` | GitHub Discussion | Read-only after publish |
| Charter, business-case, plans | Markdown in repo (no Issue) | Keep canonical in `pm-suite/projects/` |

## Operations

### push (Suite → GitHub)

For each row in the target artifact:

1. Look up (or create) the Issue via the link map.
2. Update Issue title/body to match markdown.
3. Apply labels (`pm:backlog`, `pm:risk`, `pm:decision`, `pm:benefit-Bn`, owner via assignee).
4. Add to the GitHub Project and place in the right column.

Use MCP tools (search via `ToolSearch` for these):
- `mcp__github__create_or_update_file`
- (Issue-write tooling) — check available GitHub MCP tools at runtime; `mcp__github__issue_write` for create/update, plus the Projects mutation set if exposed.

### pull (GitHub → Suite)

For each linked Issue:

1. Fetch via `mcp__github__issue_read`.
2. Update the corresponding markdown row (status, owner, comments-derived notes).
3. Surface diffs in `pm-suite/projects/<project>/.adapter-github-diff.md` for `pm-work` to resolve.

### link

`.adapter-github.json`:

```json
{
  "owner": "rickmiller434-crypto",
  "repo": "prompt-eng-interactive-tutorial",
  "project_number": 1,
  "row_links": {
    "backlog": { "OB-1": 42, "OB-2": 43 },
    "risks":   { "R1": 50, "R2": 51 },
    "decisions": { "D-001": 60 }
  }
}
```

### diff

For each linked Issue, fetch with `mcp__github__issue_read`, diff against the markdown row, and report.

## Label conventions

| Label | Meaning |
|-------|---------|
| `pm:backlog` | Backlog item |
| `pm:risk` | Risk register entry |
| `pm:decision` | Decision-log entry |
| `pm:cr` | Change request |
| `pm:benefit-B1` … `pm:benefit-Bn` | Linked benefit |
| `pm:principle-N` | Principle most relevant |
| `pm:adaptive` / `pm:predictive` | Workstream class |

## Conventions

- **Markdown is canonical** for plans, decisions, audits, charter, business case, principles.
- **GitHub is canonical** for execution state (Issue status, assignee, comments) of items the team works on there.
- Conflicts surfaced via `diff` are resolved by `pm-work` and logged in `decisions/decision-log.md`.

## Invocation

Run from a Claude Code session with the GitHub MCP connected. Use `/pm-sync github <project>` (Phase 6 command).

## Failure modes

- Rate limits → backoff + retry.
- Closed-by-bot loops (Issue closes a row but the row reopens it next sync) → adapter must respect Issue `state` if the close was annotated `pm:resolved`.
- Project schema changes (someone renames a column) → adapter halts, reports — never silently rewrites Project structure.
