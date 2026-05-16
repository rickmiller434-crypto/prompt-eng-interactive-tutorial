---
description: Run a principle audit on the project's artifacts
argument-hint: <project>
---

Run `pm-principles-auditor` on `pm-suite/projects/$ARGUMENTS/`.

Steps:

1. Confirm `tailoring.md` exists. If not, tell the user to run `/pm-init` and `/pm-tailor` first.
2. Invoke the `pm-principles-auditor` subagent via the `Agent` tool. Goal: produce a fresh audit per the agent's defined output format.
3. Save the audit to `pm-suite/projects/$ARGUMENTS/audits/audit-<YYYY-MM-DD>.md`.
4. After the audit lands, summarize the top remediations (ranked by severity) and the principle-by-principle pass/fail count. Surface any severity-4 blockers prominently.

Do NOT auto-trigger remediations. Each remediation needs a human decision on owner and timing.

If a prior audit exists, instruct the auditor to diff against it — note which findings closed, persisted, or worsened.

Cite Principle 8 (Quality).
