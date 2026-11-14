# Changelog - gcv/logger

## 0.1.0

- First version: LogDebug, LogInfo, LogWarn, LogError (adds the current VBA error), LogWrite for run-time callers.
- Line format: [source] - LEVEL: yyyy-mm-dd - hh:nn:ss message
- Daily files in %LOCALAPPDATA%\GCV\ExcelLogs\<workbook>, kept 30 days; ASCII lines (non-ASCII written as \uXXXX).
- Level per workbook: GCV_LOG_LEVEL = DEBUG in DEV; production (no name) writes INFO, WARN and ERROR.
- GcvLogExport saves the last 7 days to a file on the Desktop for support.
- Never stops the calling module; keeps the caller's Err.
