# Mission

<!--
This file is the single source of truth for the mission loop. It lives on
disk, outside the conversation transcript, so it survives auto-compaction:
every round starts by re-reading this file in full. Fill it in with
/mission:define (or by hand), then start rounds with /mission:launch.
-->

**Status:** DRAFT
<!-- DRAFT → ACTIVE → COMPLETE | STOPPED. Only /mission:round and /mission:stop change this after launch. -->

## Objective

<!-- One paragraph. What does the world look like when this mission is done?
     Write it so a stranger could judge success without asking you anything. -->

(not yet defined — run /mission:define)

## Definition of Done

<!-- A checklist of independently verifiable items. Each one must be checkable
     by inspecting the repo or running a command — "feels solid" doesn't count.
     The advisor will personally attack every checked box. -->

- [ ] (not yet defined)

## Constraints

<!-- Hard rules the work must respect: branches not to touch, APIs not to
     break, budgets, style requirements. The advisor flags any violation. -->

- Develop only on the designated feature branch; never push elsewhere.

## Non-goals

<!-- Things that look adjacent and tempting but are explicitly out of scope.
     This list is the main defense against drift. -->

- (not yet defined)

## Round log

<!-- Appended by /mission:round after every round. Never rewrite old entries —
     this is the audit trail the advisor uses to detect regressions.

     Format:
     ### Round N — <date>
     - Did: <one line>
     - Verified by: <command / evidence>
     - Advisor: NOTHING_TO_ADD | FINDINGS (top finding in one line)
     - Carry-over for next round: <one line or "none">
-->
