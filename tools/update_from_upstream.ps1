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

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo

$stamp = Get-Date -Format 'yyyy-MM-dd'
$logDir = Join-Path $repo 'logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir "upstream_update_$stamp.log"

function Write-Log([string]$msg) {
    $line = "{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $msg
    $line | Tee-Object -FilePath $log -Append
}

$py = if (Test-Path "$repo\.venv\Scripts\python.exe") { "$repo\.venv\Scripts\python.exe" } else { 'python' }
$branch = git rev-parse --abbrev-ref HEAD
$before = git rev-parse HEAD

Write-Log "=== TradingAgents upstream update ($branch @ $before) ==="

try {
    git fetch origin --tags 2>&1 | ForEach-Object { Write-Log "fetch: $_" }
    $upstream = git rev-parse origin/main
    Write-Log "origin/main = $upstream"

    # 2. mirror upstream main to the fork (never force: fork main must stay
    #    a fast-forward of upstream main, otherwise it needs manual repair).
    git push fork "${upstream}:refs/heads/main" 2>&1 | ForEach-Object { Write-Log "push main: $_" }

    # 3. rebase the integration branch onto upstream.
    $need = git rev-list "$branch..origin/main" --count
    if ([int]$need -eq 0) {
        Write-Log "branch already contains origin/main - nothing to rebase."
    }
    else {
        Write-Log "rebasing $branch onto origin/main ($need new upstream commits)."
        git rebase origin/main 2>&1 | ForEach-Object { Write-Log "rebase: $_" }
        if ($LASTEXITCODE -ne 0) {
            git rebase --abort 2>&1 | Out-Null
            Write-Log "REBASE CONFLICT - aborted automatically. Manual merge needed: git rebase origin/main"
            exit 0   # not an error: the tree is intact, just stale.
        }

        # 4. smoke-import the rebased package before keeping the result.
        & $py -c "from tradingagents.graph.trading_graph import TradingAgentsGraph; from tradingagents.default_config import build_default_config" 2>&1 |
            ForEach-Object { Write-Log "smoke: $_" }
        if ($LASTEXITCODE -ne 0) {
            git reset --hard $before 2>&1 | ForEach-Object { Write-Log "reset: $_" }
            Write-Log "SMOKE IMPORT FAILED after rebase - reset branch to $before. Manual review needed."
            exit 0
        }
        Write-Log "smoke import OK - keeping rebased branch."

        # 5. push the rebased branch.
        git push fork "$branch" --force-with-lease 2>&1 | ForEach-Object { Write-Log "push branch: $_" }
        if ($LASTEXITCODE -ne 0) {
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
