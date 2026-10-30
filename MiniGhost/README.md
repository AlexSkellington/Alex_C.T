# MiniGhost - sources

MiniGhost re-identifies a cloned TBS / SMS machine: machine name, store number, network,
INI files (Startup.ini, Server.ini, SMSStart.ini, System.ini, INFO_*901_WIN.INI, SmsHttps.INI),
the SQL tables that carry the store / terminal (TER_TAB, STO_TAB, LNK_TAB, RUN_TAB, STD_TAB),
transaction tables, GT* registry counters and old X*/Z* folders.

It ships as ONE file, `MiniGhost\MiniGhost.ps1`, because `irm https://bit.ly/MiniGhost | iex`
and the PowerShell Studio `.exe` both need a single script. This folder holds the same code split
into small files; `build\Build.ps1` puts them back together.

```
irm https://bit.ly/MiniGhost | iex
  -> MiniGhost_V38.ps1 (root launcher, never edit)
  -> https://raw.githubusercontent.com/AlexSkellington/Alex_C.T/main/MiniGhost/MiniGhost.ps1
```

Unlike the TBS script, the built `MiniGhost.ps1` **is committed**: the launcher reads it from
`main`, so **a push to `main` changes MiniGhost for everyone** the next time they run it.

## Layout

| Folder | Contents |
|---|---|
| `src\Bootstrap` | Runs first: header + version, WinForms / console helper, elevation, globals and storeman paths |
| `src\Functions\<Category>` | One function per file - `UI`, `Store`, `Sql`, `System`, `Workflow` |
| `src\Main` | `01_Load_Store_Data` (store number / name, IP, node role) and `02_Show_Form` |
| `src\GUI` | The main window and one file per button, then layout and tooltips |
| `build` | `Build.ps1`, `Manifest.psd1` (file order), `Update-FunctionIndex.ps1` |
| `tests` | `MiniGhost_SelfTest.ps1` - static checks plus the real functions run against temp files |
| `docs` | `FUNCTIONS.md` (generated: every function, its purpose and callers) |

## Change something

```powershell
# from the repository root
powershell -NoProfile -ExecutionPolicy Bypass -File .\MiniGhost\build\Build.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\MiniGhost\tests\MiniGhost_SelfTest.ps1
```

- Edit files under `src`, never `MiniGhost.ps1` (it is overwritten by the build).
- **Build.ps1** fails on parse errors, non-ASCII characters, duplicate function names and source
  files the manifest does not cover; it writes `MiniGhost.ps1` (ASCII, CRLF) and regenerates
  `docs\FUNCTIONS.md`. `-Check` exits 1 when the committed `MiniGhost.ps1` is out of date - the
  CI workflow runs it on every push.
- Commit the sources **and** the rebuilt `MiniGhost.ps1` together.
- Version bump: `$VersionNumber` / `$VersionDate` in `src\Bootstrap\01_Header_And_Version.ps1`
  plus a matching `#  vX.Y.Z  -  date` entry at the top of `CHANGELOG.txt`.
- Never run MiniGhost on a development PC to "try it": it renames the computer, rewrites the
  storeman INIs and SQL tables, and can truncate transaction tables.

The PowerShell 5.1 rules (ASCII only, no `.GetNewClosure()`, event-handler scoping, native
commands never throw, ...) are in `..\CLAUDE.md` and apply here too.
