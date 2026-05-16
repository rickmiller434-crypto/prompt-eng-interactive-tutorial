# Intake

Project intake materials. Today this is a single questionnaire — the **Tailoring Intake** — which the project owner fills out at initiation and which produces the **Tailoring Record** used by every other agent.

## Files

- `tailoring-questionnaire.md` — the questionnaire + the embedded tailoring-record output section.

## Flow

1. Copy `tailoring-questionnaire.md` into `pm-suite/projects/<project-name>/tailoring.md`.
2. Fill out Sections A–H with the project sponsor / lead.
3. Invoke `pm-tailor` (Phase 2) to populate the tailoring recommendation section.
4. Get sponsor sign-off.
5. All subsequent agent work reads this record before producing artifacts.
