# Alex_C.T - store maintenance tools

PowerShell tools for TBS / SMS store systems (Windows PowerShell 5.1).

| Tool | What it is | Run it |
|---|---|---|
| **TBS Maintenance Script** | WinForms app for the store server: server DB maintenance and backups, lane tools (SQL protocols, transactions, table pumps, time sync, DB repair), scale tools (Bizerba / Ishida, BMS, ScaleCommApp), node tools (ping, reboot, hosts, file copy) and general utilities. | `irm https://bit.ly/TBS_Maintenance_Script \| iex` (elevated PowerShell), or the packaged `TBS_Maintenance_Script_V<version>.exe` |
| **MiniGhost** | Separate single-file tool (`MiniGhost_V38.ps1`). | Run the file |

## Repository layout

```
TBS_Maintenace_Script_V2.ps1   BUILT single file (what the short link downloads) - never edit by hand
TBS_Maintenance/               sources, build, tests and docs of the TBS Maintenance Script
  src/                         one file per function + bootstrap / GUI / main-flow files
  build/                       Build.ps1, Manifest.psd1, formatter, PowerShell Studio template
  tests/                       self-test and sandboxed smoke test
  docs/                        FUNCTIONS.md (generated index), ARCHITECTURE.md
  CHANGELOG.txt
MiniGhost_V38.ps1
CLAUDE.md                      working rules for this repository (also for AI assistants)
```

The built file keeps its historical name (`Maintenace`) and path because every store
reaches it through the short link.

See [TBS_Maintenance/README.md](TBS_Maintenance/README.md) to build, test and release.
