# Architecture

How the built script runs, where its state lives and how the pieces talk. For a specific
function, look it up in [FUNCTIONS.md](FUNCTIONS.md) and open only that file.

## Startup order

The build concatenates the sources in [`build/Manifest.psd1`](../build/Manifest.psd1) order,
and the script runs top to bottom:

| Step | Files | What happens |
|---|---|---|
| 1 | `src/Bootstrap/01..06` | Launch detection (saved .ps1, packaged .exe, `irm \| iex`), version, **elevation** (relaunch with RunAs when possible), globals and paths, **storeman detection** (SMB share with a valid key -> Production; test roots -> Test; drive scan -> fallback; old key -> KEY EXPIRED), store paths |
| 2 | `src/Functions/**` | Function definitions only - nothing runs |
| 3 | `src/Init/01` | MailslotSender type |
| 4 | `src/GUI/01..14` | Builds the window in script scope: form and theme, banner, status bar, info strip, log box, idle/busy timers, protocol window, closing logic, styling helpers, the four tool buttons and menus, bottom layout |
| 5 | `src/Main/01` | **Pre-window load**, local reads only (fast): store/DB info read live from the INI files, store number, nodes from cache |
| 6 | `src/Main/02` | Registers the **warmup**: after the window is shown, network refreshes run as background jobs (nodes, lane DB info, table aliases, SMS functions), then inline steps (SQL scripts, test item, scale credentials, XE watcher, VNC export, hardware info, quarterly maintenance) |
| 7 | `src/Main/03` | `ShowDialog()` - the script waits here until the window closes, then exits |

Pre-window code must stay cache-only so the window appears quickly; anything that touches the
network belongs in the warmup or in a menu action.

## Shared state

| Variable | Holds | Written by |
|---|---|---|
| `$script:FunctionResults` | The result bag: `StoreNumber`, `StoreName`, `DBSERVER`, `DBNAME`, `ConnectionString`, `SqlModuleName`, `LaneNumToMachineName`, `ScaleCodeToIPInfo`, `BackofficeNumToMachineName`, `LaneDatabaseInfo`, `Get_Table_Aliases`, `NssmExe`, ... | Store functions, warmup jobs (merged back on the UI thread) |
| `$script:LaneProtocols` | Lane -> `TCP` / `Named Pipes` / `File` / `Offline` (also keyed by machine name) | `Start_Lane_Protocol_Jobs`, `Enable_SQL_Protocols_On_Selected_Lanes` |
| `$script:LastActivity` | Last user input, for the idle auto-close | Activity hooks (`GUI/06`, `GUI/14`) and menu clicks |
| `$script:BasePath`, `$script:OfficePath`, `$LoadPath`, ... | Storeman paths | `Bootstrap/05..06` |
| `$script:AlexRoot`, `BackupRoot`, `ScriptsFolder`, `ScriptsLog`, `SetupFiles` | Tecnica_Systems working folders | `Bootstrap/04` |

Most functions read the bag rather than taking parameters. When a function needs fresh data
it calls the producer (for example `Get_All_Lanes_Database_Info -LaneNumber 001`), which
returns from memory, then the cache file, then the network.

## Caches (`$script:SetupFiles`)

Only data that is slow to get is cached. Store info is NOT cached: it comes from local INI
files and is read live on every launch (a cached copy showed stale values until refreshed).

| File | Producer | Why it is cached |
|---|---|---|
| `SqlModule.txt` | `Get_Store_And_Database_Info` | which SQL PowerShell module exists - finding out scans every module folder; re-checked daily |
| `Retrieve_Nodes.txt` | `Retrieve_Nodes` | SQL query (or Ter_Load.sql) - lets the window open with the last node list while the fresh query runs |
| `Get_All_Lanes_Database_Info.txt` | `Get_All_Lanes_Database_Info` | per-lane DB name / server, from each lane's Startup.ini over the network (~1.5 s per lane that is off) |
| `Get_Table_Aliases.txt` | `Get_Table_Aliases` | parses every Load\*_Load.sql; validated against the SMS version and Load files |
| `Protocol_Results.txt` | lane protocol detection | network probe of every lane; scheduled tasks on the lanes read it too |

`-RefreshCache` on a producer rebuilds its cache; the warmup does that after the window shows.

## Background work

- **`Start_Background_Function_Job`** runs any script function in a background job: it ships
  the function body, `Write_Log`, any `-DependentBodies`, a copy of `FunctionResults` and
  `-SeedVars`. Only what is passed exists in the job. A shared WinForms timer (the reaper)
  collects results on the UI thread, merges the returned bag and runs the `OnComplete` step.
- **Runspace pools** for short parallel probes that must not start processes
  (`Show_Lane_SQL_VNC_Service_Statuses`, `Enable_SQL_Protocols_On_Selected_Lanes`,
  `Add_Scale_Credentials`).
- **Async .NET** for pings and DNS (`Test-Hosts_Parallel`). A PowerShell scriptblock cannot run
  on a thread-pool thread.
- **Generated scripts + scheduled tasks / NSSM services** for anything that must outlive the
  window: DB backups, SMS key sync, lane SQL enable, Remove_Archive_Bit and the duplicate-file
  monitor. Their text lives in here-strings inside the function that installs them.
- **Lane SQL protocol enable**: the settings + restart code is
  `Resources/Workers/Sql_Protocol_Settings.ps1`, embedded into both the menu action
  (`Enable_SQL_Protocols_On_Selected_Lanes`, remote registry) and the 4:15 AM task script
  (`Ensure_SQL_Enable_Scheduled_Task_For_File_Lane` + `Enable_SQL_Protocols_Worker.ps1`, local
  registry, also opens inbound TCP 1433). The task writes `<task>.result.txt` in the lane's
  Storeman\Temp; the protocol loop reads it back after the run and logs it.

## Talking to lanes and scales

Every lane action tries, in order: SQL over the lane's protocol (`$script:LaneProtocols`,
connection strings from `Get_All_Lanes_Database_Info`, System.Data.SqlClient), then a file
drop (`.sqi` / `.sql`) into the lane's `Office\XF<store><lane>` folder for SMS to process.
Programs are restarted through the `SMSStart_<store><lane>` mailslot (`MailslotSender`).
Scales are reached by IP (`ScaleCodeToIPInfo`) over `\\<ip>\c$` with the seeded `bizuser`
credentials; BMS / ScaleCommApp are driven through their config files, `DEPLOY_CHG.sql` and
trigger files in `toBizerba`.

## GUI conventions

- GUI files share script-scope variables (`$form`, `$toolTip`, `$logBox`, `$ui*` colours);
  event handlers resolve names when they fire.
- No closures: handlers that need data read it from the control's `.Tag`
  (`$New_Category_Button`, the service-status menu) or from `$script:` variables.
- Styling helpers live in `GUI/09`: `$Apply_Modern_Menu_Style`, `$Style_SubMenu`,
  `$Apply_Modern_Popup_Style`, `$New_Category_Button`.
- Each tool menu file (`GUI/10..13`) creates its items, then assembles the top level at the
  bottom: frequent actions first, the rest in submenus.
