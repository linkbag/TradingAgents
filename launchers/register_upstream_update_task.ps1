# Register the WEEKLY Sunday 22:00 "TradingAgents upstream update" task.
# ASCII only. Tries S4U first (runs even when nobody is logged in, no
# stored password). S4U self-registration can be denied without elevation;
# in that case it falls back to Interactive (runs when logged in) and says
# how to upgrade. Re-registering replaces any older task of the same name.
#
# This task runs tools/update_from_upstream.ps1: fetch TauricResearch main,
# mirror it to the linkbag fork main, rebase stockselector-integration onto
# it (auto-abort on conflict), smoke-import, push the branch. A conflicting
# or failing update leaves the working tree untouched and logs to
# logs/upstream_update_<date>.log - it never breaks live runs.

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$taskName = 'TradingAgentsUpstreamUpdate'
$action = New-ScheduledTaskAction -Execute (Join-Path $here 'update_from_upstream.bat') -WorkingDirectory $here
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At '22:00'
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew

$mode = $null
try {
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U -RunLevel Limited
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
    $mode = 'S4U (runs even when nobody is logged in)'
} catch {
    Write-Host "S4U registration denied ($($_.Exception.Message.Trim())) - falling back to Interactive."
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
    $mode = 'Interactive (runs while you are logged in). To also run when NOT logged in, right-click register_upstream_update_task.bat once and choose "Run as administrator".'
}

Write-Host "Registered $taskName : weekly Sunday 22:00."
Write-Host "Logon mode: $mode"
exit 0
