@echo off
rem -----------------------------------------------------------------
rem Weekly Sunday 22:00 TradingAgents upstream update: fetch
rem TauricResearch main, mirror to the linkbag fork, rebase the
rem stockselector-integration branch (abort-on-conflict, never breaks
rem live runs). The work happens in tools\update_from_upstream.ps1.
rem ASCII only - no Chinese in .bat files.
rem -----------------------------------------------------------------
chcp 65001>nul
cd /d %~dp0..

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\tools\update_from_upstream.ps1"
exit /b %ERRORLEVEL%
