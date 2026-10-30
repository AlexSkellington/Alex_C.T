# Alex_C.T - store maintenance tools

PowerShell tools for TBS / SMS store systems (Windows PowerShell 5.1).

| Tool | What it is | Run it |
|---|---|---|
| **TBS Maintenance Script** | WinForms app for the store server: server DB maintenance and backups, lane tools (SQL protocols, transactions, table pumps, time sync, DB repair), scale tools (Bizerba / Ishida, BMS, ScaleCommApp), node tools (ping, reboot, hosts, file copy) and general utilities. | `irm https://bit.ly/TBS_Maintenance_Script \| iex` (elevated PowerShell), or the packaged `TBS_Maintenance_Script_V<version>.exe` |
| **MiniGhost** | WinForms tool that re-identifies a cloned machine: machine name, store number, network, storeman INIs, the SQL store/terminal tables, transaction tables, GT* registry counters. | `irm https://bit.ly/MiniGhost \| iex` |

## Repository layout

```
TBS_Maintenance/        the TBS Maintenance Script - edit here
  src/                  one file per function + bootstrap / GUI / main-flow files
  build/                Build.ps1 (sources -> one script), manifest, formatter, PowerShell Studio template
  tests/                self-test and sandboxed smoke test
  docs/                 FUNCTIONS.md (generated index), ARCHITECTURE.md
  dist/                 build output (git-ignored) - published as a GitHub Release asset
  CHANGELOG.txt
MiniGhost/              MiniGhost - edit MiniGhost/src, build to MiniGhost/MiniGhost.ps1 (committed)
  src/ build/ docs/     same layout as TBS_Maintenance (see MiniGhost/README.md)
TBS_Maintenace_Script_V2.ps1   launcher for bit.ly/TBS_Maintenance_Script  - do not edit
MiniGhost_V38.ps1              launcher for bit.ly/MiniGhost               - do not edit
CLAUDE.md               working rules for this repository (also for AI assistants)
```

## Double-click launchers (Run\)

`Run\Run_MiniGhost.bat` and `Run\Run_TBS_Maintenance_Script.bat` replace the old
`Execute(...).bat` files. They download the script with Windows' `curl.exe` to `%TEMP%` and run that
file. The old ones ran `powershell -Command "irm https://bit.ly/... | iex"`, a download-and-run
command line that Microsoft Defender blocks as `Trojan:Win32/Commando.A!ml` before anything runs.

## The two root launchers

The short links cannot be changed (free bit.ly plan), so they keep pointing at these files:

| Short link | Root file (launcher) | Downloads and runs |
|---|---|---|
| `bit.ly/TBS_Maintenance_Script` | `TBS_Maintenace_Script_V2.ps1` | the latest GitHub Release asset `TBS_Maintenance_Script.ps1` |
| `bit.ly/MiniGhost` | `MiniGhost_V38.ps1` | `MiniGhost/MiniGhost.ps1` on `main` |

Each launcher is ~20 lines: download the real script, run it with `Invoke-Expression` (so it
still starts as an inline `irm | iex` launch), print a red line if the download fails. Keep
their names and paths exactly as they are.

## How the TBS script reaches stores

You edit the small files in `TBS_Maintenance/src`. `Build.ps1` joins them into one file,
`TBS_Maintenance/dist/TBS_Maintenance_Script.ps1`, because `irm | iex` and PowerShell Studio
both need a single script. That file is **not committed**: it is attached to a GitHub
Release, and the root launcher downloads the latest release:

```
https://github.com/AlexSkellington/Alex_C.T/releases/latest/download/TBS_Maintenance_Script.ps1
```

So pushing to `main` never changes what stores run - publishing a release does.
See [TBS_Maintenance/README.md](TBS_Maintenance/README.md) to build, test and release.
