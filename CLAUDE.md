# CLAUDE.md

Guidance for Claude Code (and other agents) working in this repository.

## What this repo is

A **Warhammer 40,000 mathhammer spreadsheet** for **11th edition**. It computes
expected damage and a *point efficiency* rating for every attacker/target
pairing, so you can write down a whole army's stats and compare how each unit
performs into many enemy targets at once — each pairing condensed into a single
colour-coded number.

The repository is **not a software project**: there is no source code to build,
no tests, no dependencies. The deliverable is a single binary workbook:

- **`WH40k Point Efficiency Calculator.xlsx`** — the one and only product. All
  logic lives in Excel formulas inside this file. It is ~11 MB and ~130 sheets.

Supporting files:
- `README.md` — user-facing intro and download links.
- `RELEASING.md` — **the git workflow you must follow** (see below).
- `Google Sheets version` — plain-text file holding the Google Sheets share link.
- `.gitattributes` — marks `*.xlsx`/`*.xlsm` as `binary` (no diff/merge/EOL munging).

Distribution: **git** for contributors, **Google Sheets import** for end users.

## The binary-in-git constraint (read before any commit)

The workbook is a multi-MB binary that does **not** delta-compress, so every
commit that contains it permanently adds its full size to history. The repo has
already had its history rewritten once to recover from bloat (958 MB → 33 MB).

**Committing the `.xlsx` on a per-version work branch is fine** — that's the
intended workflow. Those intermediate blobs are discarded when the branch is
collapsed, so commit (and even push as backup) freely there. What must never
happen is multiple workbook blobs landing on `main` or in a PR's history.
Follow [RELEASING.md](RELEASING.md):
- Work on a per-version branch; commit the workbook as often as you like.
- **Collapse the branch to a single commit BEFORE opening the PR**, never after
  (GitHub keeps pre-PR commits forever via `refs/pull/*`). This is the one rule
  that must never break.
- Tag each release `vMAJOR.MINOR.PATCH`; optionally attach the workbook to a
  GitHub Release so old versions stay downloadable without living in history.
- No Git LFS (deliberately declined).

Versioning: **MAJOR** = new game edition / breaking restructure, **MINOR** =
new units/factions/features, **PATCH** = data corrections & bug fixes.

`~$WH40k Point Efficiency Calculator.xlsx` is an Excel lock file — never commit it.

## Workbook architecture

Sheets run **left to right following the 40k attack sequence**, then a long tail
of hidden calculation/appendix sheets.

**Visible sheets (in order):**
1. **Intro** — overview and links.
2. **Input** — user writes their units' stats (attacks, BS/WS, S, AP, damage, special rules…).
3. **Army Rules** — programmable conditional effects for attackers (keyword/phase filtered).
4. **Targets** — defender stats (T, save, invuln, wounds, FNP, keywords, points…).
5. **Damage** — output: average damage caused, per attacker × target.
6. **Kills** — output: average models killed, per attacker × target.
7. **PPW** — output: *Points Spent Per Wound Caused*.
8. **PPP** — output: *Points Spent Per Point Killed* = the headline **point
   efficiency** rating. Also hosts the **global settings** (checkboxes; shown
   read-only on the Damage/Kills/PPW tabs).
9. **Changelog** — version history (newest at bottom).

**Hidden sheets (~120):** the actual math. They implement the attack sequence
stage by stage — hit rolls, crit/sustained/lethal, wound rolls, anti, saves
(armour/invuln/cover), AP, damage, mortal wounds, FNP, then damage
*distributions* and the dice-string "appendix" sheets (`rnd …`, `table.dmg`,
`atk.*`) that resolve random damage like `2d6+3`. Prefixes: `h` = hit,
`w` = wound, `s` = save, `d` = damage. These are intentionally hidden and are
where edition rules are encoded.

## How to inspect the workbook (you cannot run Excel here)

The `.xlsx` is a zip of XML. To read it without Excel:

```bash
unzip -o -q "WH40k Point Efficiency Calculator.xlsx" "xl/workbook.xml" -d /tmp/wb   # sheet list & order
unzip -o -q "WH40k Point Efficiency Calculator.xlsx" "xl/sharedStrings.xml" -d /tmp/wb  # all text/labels
# per-sheet cells live in xl/worksheets/sheetN.xml (map via xl/_rels/workbook.xml.rels)
```

You can read formulas and content this way, but you **cannot edit the workbook**
programmatically with confidence — formula edits, formatting, and the ~130-sheet
structure are maintained by the author in Excel. If a change to the workbook is
needed, describe exactly what to change and where; don't attempt to rewrite the
binary's XML by hand.

## Conventions & state

- Remote: `https://github.com/svknoe/WH40k-Point-Efficiency-Calculator.git`
- Default branch: `main`. Release tags: `v3.3b`, `v4.1b`, `v5.7a`, `v6.4c`, `v7.3c`, … (one kept per major).
