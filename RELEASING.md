# Release workflow

How to work on and release a new version of the calculator without bloating
the git repository. Background: the workbook is a ~12 MB binary that does not
delta-compress, so every commit containing it permanently adds its full size
to the repo history. This workflow keeps `main` at one workbook snapshot per
release.

## Version numbers

Versions follow MAJOR.MINOR.PATCH (e.g. `7.4.0`):

* **MAJOR** – new game edition or breaking restructure of the workbook
* **MINOR** – new units, factions, or features
* **PATCH** – data corrections and bug fixes

## Working on a new version

```powershell
# 1. Start a work-in-progress branch
git switch -c wip/7.4.0

# 2. Commit as often as you like. Pushing the branch as a backup is fine.
git push -u origin wip/7.4.0
```

Commit freely — the intermediate workbook blobs will be discarded later.

## Releasing

```powershell
# 3. Collapse the branch to a single commit (history rewrite)
git reset --soft main
git commit -m "Update to version 7.4.0"
git push --force origin wip/7.4.0

# 4. Open the PR and merge it (or merge locally into main and push)

# 5. Tag the release
git switch main
git pull
git tag v7.4.0
git push origin v7.4.0

# 6. Delete the WIP branch so its fat history can be garbage-collected
git branch -D wip/7.4.0
git push origin --delete wip/7.4.0

# 7. Optional but recommended: attach the workbook to a GitHub Release so
#    old versions stay downloadable without living in git history
gh release create v7.4.0 "WH40k Point Efficiency Calculator.xlsx"
```

## The one rule that must never break

**Rewrite (force-push) the branch to a single commit BEFORE opening the PR —
never after.** Once a commit has been part of a pull request, GitHub keeps it
forever via hidden `refs/pull/*` references, even if the branch is later
force-pushed or deleted. Commits that were merely pushed to a branch (no PR)
become unreachable when the branch is deleted and are garbage-collected by
GitHub within a few weeks.

## Occasional local maintenance

After deleting merged WIP branches, reclaim local disk space with:

```powershell
git gc --prune=now
```
