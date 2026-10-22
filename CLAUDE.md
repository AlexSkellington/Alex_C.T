# Working on this repository

Two PowerShell tools live here:

| Path | What it is |
|---|---|
| `TBS_Maintenance/` | **Sources** of the TBS Maintenance Script (edit these) |
| `TBS_Maintenace_Script_V2.ps1` | **Built** single-file script - generated, never edit by hand |
| `MiniGhost_V38.ps1` | MiniGhost (separate single-file tool, not part of the build) |

The built file is what ships: `irm https://bit.ly/TBS_Maintenance_Script | iex` downloads it
from `main`, and PowerShell Studio packages it into `TBS_Maintenance_Script_V<version>.exe`.
Its path and its historical `Maintenace` spelling must not change.

## Find the code first, then read only that file

1. `TBS_Maintenance/docs/FUNCTIONS.md` - every function: category, file, one-line purpose,
   and the menu item / startup step / functions that call it. Start here.
2. `TBS_Maintenance/docs/ARCHITECTURE.md` - startup order, shared state, background work,
   caches, where each GUI piece lives.
3. Open only the file you need (`src/Functions/<Category>/<Name>.ps1`). Each file is one
   function with a header block. Do not load the built script into context - it is ~35k lines.

## Make a change

```powershell
# 1. edit files under TBS_Maintenance\src
# 2. normalise whitespace, build, test
powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\build\Format-Sources.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\build\Build.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\tests\TBS_Maintenance_SelfTest.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\tests\Smoke.ps1
```

- `Build.ps1` fails on parse errors, non-ASCII characters, duplicate functions, and source
  files missing from `build/Manifest.psd1`. It also regenerates `docs/FUNCTIONS.md`.
- Commit the sources **and** the rebuilt `TBS_Maintenace_Script_V2.ps1` together.
  `Build.ps1 -Check` exits 1 when the built file is stale.
- New function: drop `src/Functions/<Category>/<Name>.ps1` (one function, same header
  format as its neighbours). The manifest picks it up; no list to edit.
- Version bump: `$VersionNumber` / `$VersionDate` in `src/Bootstrap/02_Version_And_Host.ps1`
  plus a matching `#  vX.Y.Z  -  date` entry at the top of `TBS_Maintenance/CHANGELOG.txt`
  (the self-test checks both).

## Rules that keep the stores working

- **Windows PowerShell 5.1** is the target (stores run it). No PS 7-only syntax (`??`, `?.`,
  ternary, `-Parallel`, `using namespace` mid-file).
- **ASCII only.** 5.1 reads BOM-less files as Windows-1252; one smart quote garbles the UI.
- **Tabs, Allman braces**, PowerShell Studio style - match the surrounding code.
- **Never `.GetNewClosure()`.** A closure runs in its own module scope: every `$script:`
  variable inside it is a different, empty variable. (This silently disabled lane protocol
  detection from v2.6.7 to v2.6.9.) Event handlers that need helpers: keep the helpers in
  `$script:` variables, or carry data on the control's `.Tag`.
- **`New-Object T($a + $b, $c)` is a bug** in command mode (`$a + @($b, $c)`). Use
  `[T]::new($a + $b, $c)` or parenthesise each argument.
- **`Get-Job -Name x -State y`** is an invalid parameter combination (it throws, even with
  `-ErrorAction SilentlyContinue`). Filter with `Where-Object { $_.State -eq 'Running' }`.
- **`Start-Process -ArgumentList @()`** throws on an empty array - only pass it when non-empty,
  and quote arguments that contain spaces.
- **`SqlConnectionStringBuilder`**: assign properties through `.PSBase.<Prop>` (PowerShell
  treats the builder as a dictionary; `$csb.DataSource = x` throws). Reading is fine.
- **SQL without a module**: stores may have no SqlServer/SQLPS module. New code uses
  `System.Data.SqlClient` (always present); lane connection strings carry
  `TrustServerCertificate=True` (SqlServer 22+ encrypts by default, lanes are self-signed).
- **Results leak into output**: `[System.Windows.Forms.MessageBox]::Show(...)`, `.Add()`,
  `Invoke-Sqlcmd` etc. return values - `[void]` / `$null =` them inside functions that return.
- **Arrays collapse**: `$x = $x | Where-Object ...` becomes a scalar for one item (then `+=`
  concatenates strings or throws). Wrap in `@()`. A returned DataTable is unrolled into its
  rows. `[DBNull]::Value` is truthy and `[int][DBNull]` throws.
- **Never assign automatic variables** (`$host` is read-only and throws; `$args`, `$input`,
  `$matches`, `$_` get clobbered). In `"..."` strings `$1`, `$name` expand - put literal
  dollar signs (passwords!) in single quotes or escape them.
- **schtasks / native commands**: build an argument array and call `& schtasks.exe @args`
  (never `Invoke-Expression` - it breaks on `$` and quotes); quote `/TR` paths as `\"path\"`;
  `/sc once` needs `/st`; `/ru <user>` needs `/rp` or `/it` or it prompts on the console.
  Native tools never throw: check `$LASTEXITCODE`.
- **Backups rotate after success** - never delete old backups before the new one exists.
- **Thread pool**: a scriptblock cannot run on a .NET thread-pool thread (no runspace) - use
  a runspace pool or async .NET APIs (`SendPingAsync`, `GetHostAddressesAsync`).
- **`Get-Job -Name <missing>`** records an error even with SilentlyContinue - filter with
  `Get-Job | Where-Object { $_.Name -eq ... }`.
- The built file is **LF**; newline conversions in sources use `-replace "`r?`n", "`r`n"`.
- GUI files (`src/GUI/01..14`) run top to bottom in script scope and share variables
  (`$form`, `$toolTip`, `$logBox`, `$ui*` colours, `$layout*` blocks). Handlers resolve those
  names when they fire, not when they are created.
- Functions shipped into background jobs (`Start_Background_Function_Job`, warmup steps) run
  in a fresh runspace: only what is passed in (`SeedVars`, `DependentBodies`, `FunctionResults`
  copy) exists there.
- Marker lines the tests depend on (keep them textually identical):
  - `src/Bootstrap/05_Storeman_Detection.ps1`: `$BasePath = $null` ...
    `Write-Host "Selected (storeman) folder: ...` (self-test section A/B)
  - `src/Main/01_Load_Store_Data.ps1`: the `$script:StoreNumber = [string]...` block (C)
  - `src/Bootstrap/04_Globals_And_Paths.ps1`: the `AlexRoot`, `AlexProgramDataRoot`,
    `StoremanRoots` and three `Auto*OnLaunch` lines (patched by `tests/Smoke.ps1`)
  - `function <Name>` alone on its own line (the self-test extracts functions that way)

## Testing on this PC

`tests/Smoke.ps1` runs the whole app in a sandbox: every Tecnica_Systems / ProgramData /
Storeman root is redirected to `%TEMP%\tbs-smoke`, NSSM download and the launch-time
VNC / hardware / quarterly maintenance tasks are off, the window opens off-screen and closes
itself. Never run the built script directly on a dev PC to "try it": at launch it creates
`C:\Tecnica_Systems`, downloads NSSM and can start DISM/SFC (quarterly maintenance).
Use `-Baseline` / `-Compare` to diff two smoke runs (before/after a change).

## Releasing

1. Build + both tests green, commit on a branch, merge to `main`, push.
   Stores pick up `main` immediately through the short link - push only finished work.
2. Package: open `TBS_Maintenace_Script_V2.ps1` in PowerShell Studio, Package (settings:
   `Build.ps1 -PsBuild` writes them next to the file from `build/TBS_Maintenance.psbuild.template`).
