---
description: Append a structured entry to the decision log
argument-hint: <project> "<decision text>"
---

Append a decision-log entry to `pm-suite/projects/$1/decisions/decision-log.md`.

The user has provided the decision intent in their message after the project name. Treat the remainder of `$ARGUMENTS` (or the user's full message) as the decision description.

Before writing the entry, ask the user for any missing fields:
- Rationale (1–3 sentences)
- Alternatives considered (≥ 1)
- Decision level (team / PM / sponsor / steering)
- Steward (named person)
- Affected parties
- Principles cited (at least 1)
- Linked artifacts (paths or refs)

Then append a single row to the decision log with:
- Date (today)
- ID (next sequential `D-NNN`)
- Decision (concise)
- All fields above

Conventions:
- Append at the bottom; do not edit prior entries.
- If this decision overturns a prior entry, include "Supersedes D-XXX" in the rationale.
- Tailoring changes, approved CRs, and baseline re-baselines all warrant a decision-log entry — encourage the user to run `/pm-decide` after those.

After writing, confirm to the user with the new entry's ID and the file path.

Cite Principle 1 (Stewardship), 6 (Leadership), and any others the user named.
