# TBS Maintenance Script - sources

The script ships as ONE file, `dist\TBS_Maintenance_Script.ps1` (built, not committed), so that

```powershell
irm https://bit.ly/TBS_Maintenance_Script | iex
```

and the PowerShell Studio `.exe` work - both need a single script. This folder holds the same code split into
small files that can be read and changed one at a time; `build\Build.ps1` puts them back
together.

## Layout

| Folder | Contents |
|---|---|
| `src\Bootstrap` | Runs first: launch detection, version, elevation, globals and paths, storeman detection, store paths |
| `src\Functions\<Category>` | One function per file - `Core`, `Store`, `Background`, `Dialogs`, `Server`, `Lanes`, `Scales`, `Nodes`, `General` |
| `src\Init` | Runtime init (MailslotSender type) |
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
  `-Check` exits 1 when `dist\` is out of date. Output is ASCII with LF line endings.
- **Smoke.ps1** starts the real app with every Tecnica_Systems / ProgramData / Storeman path
  redirected to `%TEMP%\tbs-smoke` and the launch-time tasks (NSSM download, VNC, hardware,
  quarterly maintenance) switched off, shows the window off-screen and reports controls,
  menus, warmup progress and every runtime error. `-Baseline <file>` / `-Compare <file>` diff
  two runs. **Do not run the built script directly on a development PC** - it creates
  `C:\Tecnica_Systems` and can start DISM/SFC.
- New function: add `src\Functions\<Category>\<Name>.ps1` with the same header format. No
  list to update.
- Commit only the sources. `dist\` is git-ignored; the built file goes out through a Release.

The rules that keep stores working (PowerShell 5.1 only, ASCII, no `.GetNewClosure()`, and the
other pitfalls found so far) are in `..\CLAUDE.md`.

## Release

Stores run the file attached to the **latest GitHub Release**. Pushing to `main` changes
nothing for them; pushing a **version tag** does: `.github/workflows/release.yml` then builds the
script on GitHub, checks the tag matches `$VersionNumber`, runs the self-test and publishes the
release with `TBS_Maintenance_Script.ps1` attached. Any failed step means no release.

1. Bump `$VersionNumber` / `$VersionDate` in `src\Bootstrap\02_Version_And_Host.ps1` and add
   the matching entry at the top of `CHANGELOG.txt` (it becomes the release notes).
2. `Build.ps1`, self-test and smoke test green locally; commit; push to `main`.
3. Tag and push the tag:
   ```powershell
   git tag v2.7.2
   git push origin v2.7.2
   ```
   Follow the run under the repository's **Actions** tab. When it is green,
   `irm https://bit.ly/TBS_Maintenance_Script | iex` runs the new version.
4. Package (optional): `Build.ps1 -PsBuild`, open `dist\TBS_Maintenance_Script.ps1` in
   PowerShell Studio and package it; the `.exe` can be uploaded to the same release.

To roll back, delete the bad release on GitHub (the previous one becomes latest again), or
release a fixed version.