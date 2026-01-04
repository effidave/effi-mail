<#
.SYNOPSIS
    Syncs git submodules - commits local changes and pulls remote updates.

.DESCRIPTION
    For each submodule:
    1. Ensures HEAD is attached to main branch
    2. Commits any uncommitted changes
    3. Pulls latest from remote
    4. Updates parent repo's submodule reference

.PARAMETER CommitMessage
    Message for any submodule commits. Default: "Update from sync script"

.EXAMPLE
    .\sync-submodules.ps1
    .\sync-submodules.ps1 -CommitMessage "Fixed typos in agent docs"
#>

param(
    [string]$CommitMessage = "Update from sync script"
)

$ErrorActionPreference = "Stop"
$RepoRoot = git rev-parse --show-toplevel
Set-Location $RepoRoot

Write-Host "`n=== Syncing Submodules ===" -ForegroundColor Cyan
Write-Host "Repo root: $RepoRoot" -ForegroundColor Gray

# Get list of submodules
$gitmodulesPath = Join-Path $RepoRoot ".gitmodules"
if (-not (Test-Path $gitmodulesPath)) {
    Write-Host "No .gitmodules file found." -ForegroundColor Yellow
    exit 0
}

$submodules = git config --file $gitmodulesPath --get-regexp path | ForEach-Object {
    ($_ -split ' ')[1]
}

if (-not $submodules) {
    Write-Host "No submodules found." -ForegroundColor Yellow
    exit 0
}

$updatedSubmodules = @()

foreach ($submodule in $submodules) {
    $submodulePath = Join-Path $RepoRoot $submodule
    
    Write-Host "`n--- $submodule ---" -ForegroundColor Yellow
    
    Push-Location $submodulePath
    try {
        # Check current state
        $branch = git rev-parse --abbrev-ref HEAD 2>$null
        $isDirty = (git status --porcelain) -ne $null
        
        # Reattach to main if detached
        if ($branch -eq "HEAD") {
            Write-Host "  Reattaching to main branch..." -ForegroundColor Gray
            $ErrorActionPreference = "Continue"
            $null = git checkout main 2>&1
            if ($LASTEXITCODE -ne 0) {
                # main doesn't exist locally, create it tracking origin/main
                $null = git checkout -b main origin/main 2>&1
            }
            $ErrorActionPreference = "Stop"
        }
        
        # Commit local changes if any
        if ($isDirty) {
            Write-Host "  Committing local changes..." -ForegroundColor Gray
            git add -A
            git commit -m $CommitMessage
            $updatedSubmodules += $submodule
        }
        
        # Pull latest from remote
        Write-Host "  Pulling from remote..." -ForegroundColor Gray
        $ErrorActionPreference = "Continue"
        $pullResult = git pull --rebase 2>&1 | Out-String
        $ErrorActionPreference = "Stop"
        if ($pullResult -match "Already up to date") {
            Write-Host "  Already up to date." -ForegroundColor Green
        } else {
            Write-Host "  Pulled updates." -ForegroundColor Green
            $updatedSubmodules += $submodule
        }
        
        # Push if we have local commits ahead of origin
        $ahead = git rev-list --count origin/main..HEAD 2>$null
        if ($ahead -gt 0) {
            Write-Host "  Pushing $ahead commit(s) to remote..." -ForegroundColor Gray
            $ErrorActionPreference = "Continue"
            git push origin main 2>&1 | Out-Null
            $ErrorActionPreference = "Stop"
            Write-Host "  Pushed." -ForegroundColor Green
        }
        
    } finally {
        Pop-Location
    }
}

# Update parent repo's submodule references
$updatedSubmodules = $updatedSubmodules | Select-Object -Unique
if ($updatedSubmodules.Count -gt 0) {
    Write-Host "`n--- Parent Repo ---" -ForegroundColor Yellow
    Write-Host "  Updating submodule references..." -ForegroundColor Gray
    
    foreach ($sub in $updatedSubmodules) {
        git add $sub
    }
    
    $subList = $updatedSubmodules -join ", "
    git commit -m "Update submodule(s): $subList"
    
    Write-Host "  Pushing parent repo..." -ForegroundColor Gray
    git push
    Write-Host "  Done." -ForegroundColor Green
}

Write-Host "`n=== Sync Complete ===" -ForegroundColor Cyan
