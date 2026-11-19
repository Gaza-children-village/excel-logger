# GCV Logger

![Release](https://img.shields.io/github/v/release/Gaza-children-village/excel-logger?filter=logger-v*&label=release)
![Excel](https://img.shields.io/badge/Excel-2016%2B%20Windows-217346)
![VBA](https://img.shields.io/badge/VBA-add--in-blue)

Excel add-in (`GCV Logger.xlam`) that keeps the logs of every GCV add-in and lets a secretary send them to the technical manager in two clicks.

## Ribbon

| Button | What it does |
|---|---|
| **عرض السجلات** (show logs) | Opens the log folder with today's file selected |
| **الإبلاغ عن مشكلة** (report an issue) | Choose all add-ins (default) or one, add a short description; a zip with `report.txt` (versions, Excel, Windows, description) and the last 14 days of logs is saved on the Desktop and an Outlook e-mail opens with it attached |

## What is logged

What the add-ins **do**, never the data they handle: no names, no ID numbers, no other personal data.

```
[student-lookup] - INFO: 2026-11-19 - 09:30:12 Lookup started: 3 row(s).
[student-lookup] - DEBUG: 2026-11-19 - 09:30:12 Row 4: found
[updater] - WARN: 2026-11-19 - 09:31:02 No internet connection; nothing checked.
```

- Files: `%LOCALAPPDATA%\GCV\ExcelLogs\GCV Tools\yyyy-mm-dd.log`, one per day, kept 30 days.
- Levels: DEBUG, INFO, WARN, ERROR. Production PCs write INFO and above; a development PC can switch to DEBUG with `GcvLogUseDebugLevel` (stored per user in the registry).
- Logging never stops the add-in that logs, and leaves its `Err` unchanged.

## For other add-ins

Other GCV add-ins include `GcvLogClient.bas` and call `LogInfo "my-addin", "message"` (also `LogDebug`, `LogWarn`, `LogError`). The client forwards the line with `Application.Run` to this add-in, or to the Immediate window when it is not loaded. No reference between projects is needed.

## Settings

`SUPPORT_EMAIL` in `src/LoggerUi.bas`: the address the report e-mail is addressed to (empty: the secretary types it).

## Installing

Secretaries do not install add-ins one by one. The **GCV Tools installer** (see [excel-updater](https://github.com/Gaza-children-village/excel-updater#installing-on-a-pc)) installs all of them for the current Windows user, and the **update tools** button keeps them current.

## Building from source

Excel cannot load a VBA project created outside Excel, so the add-in is assembled in Excel from the files in this repository:

1. Copy `GCV Logger - shell.xlam` and rename the copy to `GCV Logger.xlam` (the shell already contains the add-in file type and the ribbon).
2. Open it, press **Alt+F11** and select its VBAProject.
3. **File › Import File**: every file in `src/`.
4. Double-click **ThisWorkbook** and paste the contents of `ThisWorkbook.code`.
5. **Debug › Compile VBAProject**, then save with **Ctrl+S**. Saving records the add-in's id, version and file name in the file.

Source files are ASCII with CRLF line endings. Arabic text is stored as Unicode code points (`FromCodePoints("0627 ...")`), because the VBA editor cannot keep Arabic characters.

## Tests

Run `RunLoggerTests` and `RunLoggerReportTests` from the VBA editor (click inside the procedure and press **F5**). Results appear in a message box and in the Immediate window (**Ctrl+G**). The tests use fakes and temporary files: no network, no real data.

Before a release, remove the development modules (`TestGcvLog`, `TestLoggerReport`), compile and save.

## Releasing

Releases are checked and published by GitHub Actions ([`.github/workflows/ci.yml`](.github/workflows/ci.yml)); only the build happens by hand, because it needs Excel.

1. Raise `ADDIN_VERSION` in `src/GcvAddin.bas` and add a `## <version>` section to `CHANGELOG.md`.
2. Build the add-in in Excel (see above) **without** the test modules, and save it.
3. Upload it as a **draft** release:
   ```
   gh release create logger-v<version> "GCV Logger.xlam" --draft --repo Gaza-children-village/excel-logger
   ```
4. Merge to `main` and push the tag `logger-v<version>` (through Gitea; the mirror carries it to GitHub).
5. The workflow checks the uploaded file (same id and version as the tag, add-in type, ribbon, no test modules, CHANGELOG section), adds `manifest.json` (SHA-256 and size), renames the file to its asset name, sets the release notes from `CHANGELOG.md` and **publishes**. If anything is wrong, the release stays a draft (the updater ignores drafts) and the job shows why.

Uploaded the draft after pushing the tag? Re-run the failed **release** job in the Actions tab, or `gh workflow run ci.yml -f tag=logger-v<version> --repo Gaza-children-village/excel-logger`.

Every push also runs the source checks: VBA files ASCII with CRLF, module names equal to file names, `Option Explicit`, the shared ribbon tab, a valid identity in `GcvAddin.bas` with a CHANGELOG section. Run them locally with `python .github/scripts/gcv_ci.py check`.

## Files

| File | Purpose |
|---|---|
| `src/GcvLog.bas` | The logger: levels, files, retention, `LogWrite` |
| `src/LoggerUi.bas` | Ribbon, show logs, issue report (zip, e-mail) |
| `src/GcvAddin.bas` | Identity of the add-in (id, version, repository) |
| `src/Test*.bas` | Tests (development only) |
| `customUI14.xml` | Ribbon definition (already inside the shell) |

## Part of GCV Tools

Excel add-ins for the schools of Gaza Children's Village. They share one ribbon tab, **Gaza Children Village Tools**, and are installed and updated together.

| Add-in | Repository | What it does |
|---|---|---|
| GCV Logger | [excel-logger](https://github.com/Gaza-children-village/excel-logger) | Logs of all GCV add-ins; show logs, report an issue |
| GCV Updater | [excel-updater](https://github.com/Gaza-children-village/excel-updater) | Installs updates and new tools; home of the catalog, loader and installer |
| GCV Student Lookup | [excel-student-lookup](https://github.com/Gaza-children-village/excel-student-lookup) | Fills student rows from the Civil Registry, on demand |
| GCV English Names | [excel-english-names](https://github.com/Gaza-children-village/excel-english-names) | `=TO_ENGLISH(cell)`: English spelling of Arabic names |

See [CHANGELOG.md](CHANGELOG.md).
