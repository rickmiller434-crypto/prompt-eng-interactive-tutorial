# Slash Commands

User-facing `/pm-*` slash commands for Claude Code. Each command is a thin prompt that invokes one or more agents from the suite.

## Roster

| Command | Argument hint | What it does |
|---------|----------------|--------------|
| `/pm-init` | `<project-name>` | Initialize a new project directory and start tailoring intake |
| `/pm-tailor` | `<project>` | Invoke `pm-tailor` on the project's filled intake |
| `/pm-status` | `<project>` | Generate a status report at the project's cadence |
| `/pm-audit` | `<project>` | Run `pm-principles-auditor` |
| `/pm-decide` | `<project>` `"<decision text>"` | Append a structured decision to the decision log |
| `/pm-risk` | `<project>` | Open the risk register for a review pass |
| `/pm-retro` | `<project>` `<sprint #>` | Run a retrospective |
| `/pm-cr` | `<project>` `"<change request title>"` | Open a new change request |
| `/pm-close` | `<project>` | Run the closure workflow (lessons learned → benefits realization → handover → closure report) |
| `/pm-sync` | `notion\|github` `<project>` | Run a storage adapter round-trip |

## Installation

```bash
mkdir -p .claude/commands
cp pm-suite/commands/*.md .claude/commands/
```

After that, the commands are callable from any Claude Code session in this repo.

## Conventions

- Every command reads `pm-suite/projects/<project>/tailoring.md` before acting (when a project exists).
- Every command produces or updates a single artifact, then halts — no chained side effects without explicit user confirmation.
- Commands cite principles in their output, matching the agent they invoke.
