# Alex_C.T - store maintenance tools

PowerShell tools for TBS / SMS store systems (Windows PowerShell 5.1).

| Tool | What it is | Run it |
|---|---|---|
| **TBS Maintenance Script** | WinForms app for the store server: server DB maintenance and backups, lane tools (SQL protocols, transactions, table pumps, time sync, DB repair), scale tools (Bizerba / Ishida, BMS, ScaleCommApp), node tools (ping, reboot, hosts, file copy) and general utilities. | `irm https://bit.ly/TBS_Maintenance_Script \| iex` (elevated PowerShell), or the packaged `TBS_Maintenance_Script_V<version>.exe` |
| **MiniGhost** | Separate tool (`MiniGhost_V38.ps1`). | `irm https://bit.ly/MiniGhost \| iex` |

## Repository layout

```
TBS_Maintenance/        the TBS Maintenance Script - edit here
  src/                  one file per function + bootstrap / GUI / main-flow files
  build/                Build.ps1 (sources -> one script), manifest, formatter, PowerShell Studio template
  tests/                self-test and sandboxed smoke test
  docs/                 FUNCTIONS.md (generated index), ARCHITECTURE.md
  dist/                 build output (git-ignored) - published as a GitHub Release asset
  CHANGELOG.txt
MiniGhost_V38.ps1       MiniGhost (separate project)
CLAUDE.md               working rules for this repository (also for AI assistants)
```

## How the TBS script reaches stores

You edit the small files in `TBS_Maintenance/src`. `Build.ps1` joins them into one file,
`TBS_Maintenance/dist/TBS_Maintenance_Script.ps1`, because `irm | iex` and PowerShell Studio
both need a single script. That file is **not committed**: it is attached to a GitHub
Release, and the short link points at the latest release:

```
https://github.com/AlexSkellington/Alex_C.T/releases/latest/download/TBS_Maintenance_Script.ps1
```

So pushing to `main` never changes what stores run - publishing a release does.
See [TBS_Maintenance/README.md](TBS_Maintenance/README.md) to build, test and release.
