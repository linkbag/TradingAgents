@echo off
rem -----------------------------------------------------------------
rem Register the WEEKLY Sunday 22:00 "TradingAgents upstream update"
rem task. The work happens in register_upstream_update_task.ps1; S4U
rem is preferred and an Interactive fallback is automatic when
rem elevation is required. Re-running replaces any earlier task of
rem the same name.
rem ASCII only - no Chinese in .bat files.
rem -----------------------------------------------------------------
chcp 65001>nul
cd /d %~dp0
set "TASK_NAME=TradingAgentsUpstreamUpdate"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0register_upstream_update_task.ps1"
if errorlevel 1 (
    echo ERROR: failed to register scheduled task %TASK_NAME%.
    exit /b 1
)
schtasks /query /tn %TASK_NAME% /fo LIST
