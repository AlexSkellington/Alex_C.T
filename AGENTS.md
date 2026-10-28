# Agent guide for this repository

The rules for working in this repo are in [`CLAUDE.md`](CLAUDE.md) - they apply to every agent,
not only Claude. Read it fully before editing: it explains the source layout, the build, the
tests, the release and the Windows PowerShell 5.1 pitfalls that already broke stores.

Then:

- Find a function in [`TBS_Maintenance/docs/FUNCTIONS.md`](TBS_Maintenance/docs/FUNCTIONS.md)
  and open only that file under `TBS_Maintenance/src/`.
- How the pieces fit: [`TBS_Maintenance/docs/ARCHITECTURE.md`](TBS_Maintenance/docs/ARCHITECTURE.md).
- Never edit `TBS_Maintenance/dist/` (build output); never move or rename the two root
  launchers (`TBS_Maintenace_Script_V2.ps1`, `MiniGhost_V38.ps1`).
- A pushed version tag releases to the stores; a push to `main` changes MiniGhost for its
  users. Do neither without the owner's go-ahead.

If this checkout sits in the owner's shared workspace (`..\AGENTS.md` and `..\CHANGES.md`
exist next to it), read those too and add an entry at the top of `..\CHANGES.md` after every
change.
