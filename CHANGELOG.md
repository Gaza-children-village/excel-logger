# Changelog - gcv/logger (GCV Logger add-in)

## 1.0.0

- Separate Excel add-in "GCV Logger.xlam"; every GCV add-in logs through it.
- Ribbon: "show logs" opens the log folder; "report an issue" zips report.txt (versions, Excel, Windows, description) and the last 14 days of logs - for all add-ins (default) or one - to the Desktop and opens an Outlook e-mail with the zip.
- One log folder for all add-ins: %LOCALAPPDATA%\GCV\ExcelLogs\GCV Tools; level in the registry (DEBUG on DEV PCs).
