param(
    [Parameter(Mandatory=$true)]
    [string]$RepoUrl,

    [Parameter(Mandatory=$true)]
    [string]$OutputPath,

    [Parameter(Mandatory=$false)]
    [string]$PrimaryBranchName = "main",

    [Parameter(Mandatory=$false)]
    [string]$git = "git"
)

if (!(Test-Path $OutputPath)) {
    Write-Host "Creating output directory: $OutputPath" -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

Push-Location $OutputPath

$env:GIT_CONFIG_GLOBAL = Join-Path -path (pwd) -childpath ".gitconfig"
# Default Git settings
& $git config --global user.email "tfs@git"
& $git config --global user.name "TFS migration"

& $git config --global core.autocrlf false
& $git config --global core.longpaths true
# Old TFS checkins are case-insensitive, so we need to ignore case.
& $git config --global core.ignorecase true

# Disable special unicode file name treatments
& $git config --global core.quotepath false

& $git config --global --add safe.directory '*'

# Rename normal .gitignore (incase that is part of existing tfs structure)
& $git config --global core.excludesFile .tfs-gitignore


# Clone the main repo
Write-Host "Cloning repository to: $PrimaryBranchName" -ForegroundColor Cyan
& $git clone -b $PrimaryBranchName $RepoUrl $PrimaryBranchName

Push-Location $PrimaryBranchName

# Get all remote branches (excluding HEAD)
$Branches = & $git branch -r | Where-Object { $_ -notmatch "HEAD" } | ForEach-Object { $_.Trim() -replace "origin/", "" }

Write-Host "Found branches: $($Branches -join ', ')" -ForegroundColor Yellow

# Create worktrees for each branch
foreach ($Branch in $Branches) {
    if ($Branch -eq $PrimaryBranchName) {
        continue 
    }
    Write-Host "Creating worktree for branch: $Branch" -ForegroundColor Green
    & $git worktree add "..\$Branch" $Branch
}

# List all created worktrees
Write-Host "`nCreated worktrees:" -ForegroundColor Yellow
& $git worktree list

Pop-Location # $PrimaryBranchName
Pop-Location # $OutputPath

Write-Host "`nScript completed! All branches are now available in: $OutputPath" -ForegroundColor Green

Write-Host "Before continuing with convertion, place the laststate.json file in the output folder $OutputPath"
