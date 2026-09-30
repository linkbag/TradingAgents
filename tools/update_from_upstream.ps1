# update_from_upstream.ps1 - weekly TradingAgents upstream sync
#
# Runs unattended (scheduled Sunday 22:00). In order:
#   1. git fetch origin (TauricResearch/TradingAgents)
#   2. mirror origin/main -> fork main (github.com/linkbag/TradingAgents)
#   3. rebase the CURRENT branch (stockselector-integration) onto origin/main
#      - a conflicting rebase is aborted automatically; the tree is untouched
#        and the log asks for a manual merge
#   4. smoke-import the rebased package with the repo venv; on failure the
#      branch is reset to its pre-rebase commit (recorded before rebase)
#   5. push the (possibly updated) branch to the fork
#
# Log: logs/upstream_update_YYYY-MM-DD.log (one file per run, never deleted).

$ErrorActionPreference = 'Continue'   # native git stderr must not become terminating (PS 5.1)
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo

$stamp = Get-Date -Format 'yyyy-MM-dd'
$logDir = Join-Path $repo 'logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir "upstream_update_$stamp.log"

function Write-Log([string]$msg) {
    # Add-Content + Write-Host: deliberately no pipeline output, so callers
    # can use Write-Log without polluting return values.
    Add-Content -Path $log -Value ("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $msg)
    Write-Host ("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $msg)
}

# Run a native command, append its output to the log, return ONLY the exit code.
function Invoke-Logged([string]$tool, [string[]]$argv) {
    $out = & $tool @argv 2>&1 | ForEach-Object { "$_" }
    foreach ($line in $out) { if ($line -match '\S') { Write-Log $line } }
    return $LASTEXITCODE
}

$py = if (Test-Path "$repo\.venv\Scripts\python.exe") { "$repo\.venv\Scripts\python.exe" } else { 'python' }
$branch = git rev-parse --abbrev-ref HEAD
$before = git rev-parse HEAD

Write-Log "=== TradingAgents upstream update ($branch @ $before) ==="

try {
    $rc = Invoke-Logged git @('fetch', 'origin', '--tags')
    if ($rc -ne 0) { Write-Log "FETCH FAILED (exit $rc) - continuing with last known origin/main." }
    $upstream = git rev-parse origin/main
    Write-Log "origin/main = $upstream"

    # 2. mirror upstream main to the fork (never force: fork main must stay
    #    a fast-forward of upstream main, otherwise it needs manual repair).
    $rc = Invoke-Logged git @('push', 'fork', "${upstream}:refs/heads/main")
    if ($rc -ne 0) { Write-Log "FORK MAIN PUSH FAILED (exit $rc) - manual repair needed." }

    # 3. rebase the integration branch onto upstream.
    $need = git rev-list "$branch..origin/main" --count
    if ([int]$need -eq 0) {
        Write-Log "branch already contains origin/main - nothing to rebase."
    }
    else {
        Write-Log "rebasing $branch onto origin/main ($need new upstream commits)."
        $rc = Invoke-Logged git @('rebase', 'origin/main')
        if ($rc -ne 0) {
            Invoke-Logged git @('rebase', '--abort') | Out-Null
            Write-Log "REBASE CONFLICT - aborted automatically. Manual merge needed: git rebase origin/main"
            exit 0   # not an error: the tree is intact, just stale.
        }

        # 4. smoke-import the rebased package before keeping the result.
        $rc = Invoke-Logged $py @('-c', 'from tradingagents.graph.trading_graph import TradingAgentsGraph; from tradingagents.default_config import build_default_config')
        if ($rc -ne 0) {
            Invoke-Logged git @('reset', '--hard', $before) | Out-Null
            Write-Log "SMOKE IMPORT FAILED after rebase - reset branch to $before. Manual review needed."
            exit 0
        }
        Write-Log "smoke import OK - keeping rebased branch."

        # 5. push the rebased branch.
        $rc = Invoke-Logged git @('push', 'fork', $branch, '--force-with-lease')
        if ($rc -ne 0) {
            Write-Log "branch push FAILED (rebased result is local-only). Manual push needed."
            exit 0
        }
    }

    Write-Log "=== done ==="
    exit 0
}
catch {
    Write-Log "UNEXPECTED ERROR: $_"
    exit 1
}
