# Install — User-Level Claude Code Configuration (Rick Miller)

This installs the Ausenco FD Presentation Standard at the **user level**
on your local Windows machine, so it auto-activates on **every Claude Code
session, in every workspace**, not just one project.

## What gets installed where

The files in this `.claude-user/` directory mirror the structure they should
have at `C:\Users\rickm\.claude\` on your Windows machine.

| Source (this repo) | Destination (your machine) |
|---|---|
| `.claude-user/CLAUDE.md` | `C:\Users\rickm\.claude\CLAUDE.md` |
| `.claude-user/skills/ausenco-fd-presentation-standard.md` | `C:\Users\rickm\.claude\skills\ausenco-fd-presentation-standard.md` |
| `.claude-user/templates/ausenco/template_memo.md` | `C:\Users\rickm\.claude\templates\ausenco\template_memo.md` |
| `.claude-user/templates/ausenco/template_workbook_cover.xlsx` | `C:\Users\rickm\.claude\templates\ausenco\template_workbook_cover.xlsx` |
| `.claude-user/templates/ausenco/template_audit_checklist.md` | `C:\Users\rickm\.claude\templates\ausenco\template_audit_checklist.md` |
| `.claude-user/scripts/fd_audit.py` | `C:\Users\rickm\.claude\scripts\fd_audit.py` |

## Install steps (Windows PowerShell)

Run this from PowerShell after cloning or pulling this repo:

```powershell
# 1. Clone or pull this repo to wherever you keep code, e.g. C:\code\
cd C:\code\
git clone https://github.com/rickmiller434-crypto/prompt-eng-interactive-tutorial.git
# or if already cloned:
cd prompt-eng-interactive-tutorial
git pull

# 2. Make sure ~/.claude directories exist
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.claude" | Out-Null
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.claude\skills" | Out-Null
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.claude\templates\ausenco" | Out-Null
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.claude\scripts" | Out-Null

# 3. Copy from this repo's .claude-user/ to your user-level ~/.claude/
$src = ".\.claude-user"
$dst = "$env:USERPROFILE\.claude"

Copy-Item -Force "$src\CLAUDE.md" "$dst\CLAUDE.md"
Copy-Item -Force "$src\skills\ausenco-fd-presentation-standard.md" "$dst\skills\"
Copy-Item -Force "$src\templates\ausenco\*" "$dst\templates\ausenco\"
Copy-Item -Force "$src\scripts\fd_audit.py" "$dst\scripts\"

# 4. Verify
Get-ChildItem -Recurse "$env:USERPROFILE\.claude" | Select-Object FullName, Length
```

## Install steps (Mac / Linux)

```bash
cd ~/code/
git clone https://github.com/rickmiller434-crypto/prompt-eng-interactive-tutorial.git
cd prompt-eng-interactive-tutorial

mkdir -p ~/.claude/{skills,scripts,templates/ausenco}
cp .claude-user/CLAUDE.md                                  ~/.claude/
cp .claude-user/skills/*                                   ~/.claude/skills/
cp .claude-user/templates/ausenco/*                        ~/.claude/templates/ausenco/
cp .claude-user/scripts/fd_audit.py                        ~/.claude/scripts/

ls -la ~/.claude/
```

## How to verify it's working

Open Claude Code in any new workspace and ask:

> "What are my active engagements?"

If installed correctly, Claude should answer with the list from
`~/.claude/CLAUDE.md` (Goldboro / Green Bay / Goderich / Kemess) without
you loading any file. If Claude doesn't know, the user-level CLAUDE.md
isn't being auto-loaded — re-check the install path and that the file is
literally at `C:\Users\rickm\.claude\CLAUDE.md`.

## What CHANGES vs the previous project-level install

| Before | After |
|---|---|
| `<workspace>/CLAUDE.md` — fires only when this repo is open | `~/.claude/CLAUDE.md` — fires on every Claude Code session, every workspace |
| `<workspace>/.claude/skills/ausenco-fd-presentation-standard.md` | `~/.claude/skills/ausenco-fd-presentation-standard.md` |
| Templates in `<workspace>/mining_estimating_manual/templates/` | `~/.claude/templates/ausenco/` |
| Audit at `<workspace>/mining_estimating_manual/build/fd_audit.py` | `~/.claude/scripts/fd_audit.py` |

The project-level files in this repo still work for THIS workspace and are
kept as backup. But the user-level install is the durable layer that
applies across all your Claude Code work — Carter's workspace, Kemess
workspace, BD pursuit workspace, future client repos, anywhere.

## Updating

When the SKILL is revised (e.g. you absorb another lesson from a parallel
deliverable):

1. Edit the source file in this repo's `.claude-user/skills/` directory.
2. Commit + push.
3. On your Windows machine: pull the repo, re-run the copy command from
   step 3 of the install. ~30 seconds.

For convenience you can save the PowerShell copy block above as
`C:\code\prompt-eng-interactive-tutorial\update-claude-config.ps1` and run it
whenever you pull the repo.

## Honest caveats

1. **Claude Code on the web (browser, like this session) — the install is
   irrelevant.** Web sessions run in ephemeral remote containers; they don't
   read your local `C:\Users\rickm\.claude\`. The install only matters for
   Claude Code CLI running on your actual Windows machine.

2. **claude.ai web chat — the install is also irrelevant.** Web chat doesn't
   read your local files. For claude.ai durability, use the "Projects"
   feature and upload the SKILL as a project file. That's a separate
   mechanism.

3. **Different machines = re-install.** If you use Claude Code on multiple
   computers, run the install on each.
