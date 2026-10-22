# TBS Maintenance Script - sources

The script ships as ONE file, `..\TBS_Maintenace_Script_V2.ps1`, so that

```powershell
irm https://bit.ly/TBSMaintenanceScript | iex
```

and the PowerShell Studio `.exe` keep working. This folder holds the same code split into
small files that can be read and changed one at a time; `build\Build.ps1` puts them back
together.

## Layout

| Folder | Contents |
|---|---|
| `src\Bootstrap` | Runs first: launch detection, version, elevation, globals and paths, storeman detection, store paths |
| `src\Functions\<Category>` | One function per file - `Core`, `Store`, `Background`, `Dialogs`, `Server`, `Lanes`, `Scales`, `Nodes`, `General` |
| `src\Init` | Runtime init and the headless `-LaneSqlEnableWorker` mode |
| `src\GUI` | The main window, top to bottom: form/theme, banner, status bar, info strip, log, idle/busy, protocol window, closing, styling helpers, the four tool menus, layout |
| `src\Main` | Load store data, startup warmup, show the window |
| `src\Resources` | Embedded files (C# renderer, worker scripts) - pulled in by `#TBS:EMBED` lines |
| `build` | `Build.ps1`, `Manifest.psd1` (file order), `Format-Sources.ps1`, `Update-FunctionIndex.ps1`, PowerShell Studio template |
| `tests` | `TBS_Maintenance_SelfTest.ps1` (85 checks), `Smoke.ps1` (whole app in a sandbox) |
| `docs` | `FUNCTIONS.md` (generated: every function, its file, purpose and callers), `ARCHITECTURE.md` |

Each function file starts with a header block: description, parameters, the shared variables
it reads/writes, what it writes on disk. Start at `docs\FUNCTIONS.md` to find the file.

## Change something

```powershell
# from the repository root
powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\build\Format-Sources.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\build\Build.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\tests\TBS_Maintenance_SelfTest.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\TBS_Maintenance\tests\Smoke.ps1
```

- **Build.ps1** fails on parse errors, non-ASCII characters, duplicate function names and
  source files the manifest does not cover; it regenerates `docs\FUNCTIONS.md`.
  `-Check` exits 1 when the built file is out of date. Output is ASCII with LF line endings.
- **Smoke.ps1** starts the real app with every Tecnica_Systems / ProgramData / Storeman path
  redirected to `%TEMP%\tbs-smoke` and the launch-time tasks (NSSM download, VNC, hardware,
  quarterly maintenance) switched off, shows the window off-screen and reports controls,
  menus, warmup progress and every runtime error. `-Baseline <file>` / `-Compare <file>` diff
  two runs. **Do not run the built script directly on a development PC** - it creates
  `C:\Tecnica_Systems` and can start DISM/SFC.
- New function: add `src\Functions\<Category>\<Name>.ps1` with the same header format. No
  list to update.
- Commit the sources **and** the rebuilt `TBS_Maintenace_Script_V2.ps1` together.

The rules that keep stores working (PowerShell 5.1 only, ASCII, no `.GetNewClosure()`, and the
other pitfalls found so far) are in `..\CLAUDE.md`.

## Release

1. Bump `$VersionNumber` / `$VersionDate` in `src\Bootstrap\02_Version_And_Host.ps1` and add
   the matching entry at the top of `CHANGELOG.txt` (the self-test checks both).
2. Build, self-test and smoke test green; commit on a branch; merge to `main`; push.
   **Stores pick up `main` immediately** through the short link.
3. Package: `Build.ps1 -PsBuild` writes the PowerShell Studio settings next to the built file;
   open `TBS_Maintenace_Script_V2.ps1` in PowerShell Studio and package it.
