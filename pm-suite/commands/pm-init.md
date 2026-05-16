---
description: Initialize a new project in the PM Suite and start the tailoring intake
argument-hint: <project-name>
---

You are initializing a new project named `$ARGUMENTS` in the PM Suite.

Steps:

1. Create `pm-suite/projects/$ARGUMENTS/` and the standard subdirectories:
   - `decisions/`
   - `status-reports/`
   - `audits/`
   - `sprint-records/` (if adaptive or hybrid lifecycle anticipated)
   - `closure/`

2. Copy `pm-suite/intake/tailoring-questionnaire.md` to `pm-suite/projects/$ARGUMENTS/tailoring.md`.

3. Replace the `<project name>` placeholders with `$ARGUMENTS`.

4. Ask the user to fill out Sections A–H of the tailoring intake. Tell them they can paste answers in a single message; do not block on a perfect fill — Section A and the sponsor's identity are the minimum.

5. After the user provides answers, save them into `tailoring.md` and tell them the next step is `/pm-tailor $ARGUMENTS` to produce the Tailoring Record.

Do NOT invoke `pm-tailor` automatically. The user controls when to advance.

Cite principles in your reply: 7 (Tailoring), and any others that apply to the project's purpose as the user describes it.
