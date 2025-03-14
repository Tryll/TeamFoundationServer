<#
.SYNOPSIS
    Converts a TFVC repository to Git while preserving complete history.

.DESCRIPTION
    ConvertTo-Git.ps1 extracts TFVC history and directly replays/applies it to a Git repository.
    This script processes all changesets chronologically through the entire branch hierarchy,
    maintaining original timestamps, authors, and comments. It creates a consistent, flat 
    migration suitable for large projects with complex branch structures.

.PARAMETER TfsPath
    The TFVC path to convert, in the format "$/ProjectName".
    This can be the root of a project or a subfolder.

.PARAMETER TfsCollection
    The URL to your TFS/Azure DevOps collection, e.g., "https://Some.Private.Server/tfs/DefaultCollection".

.PARAMETER OutputPath
    The local folder where the Git repository will be created/updated.

.PARAMETER TfsUserName
    Optional username for TFS/Azure DevOps authentication.
    If not provided, Windows integrated authentication will be used.

.PARAMETER TfsPassword
    Optional password for TFS/Azure DevOps authentication.
    If not provided but TfsUserName is, the script will either:
    - Check for an environment variable named "TfsPassword"
    - Prompt for the password interactively using a secure prompt

.EXAMPLE
    # Using Windows Authentication:
    .\ConvertTo-Git.ps1 -TfsPath "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://Some.Private.Server/tfs/DefaultCollection"

.EXAMPLE
    # Using username with interactive password prompt:
    .\ConvertTo-Git.ps1 -TfsPath "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://Some.Private.Server/tfs/DefaultCollection" -TfsUserName "your_username"

.EXAMPLE
    # Using username and password:
    .\ConvertTo-Git.ps1 -TfsPath "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://Some.Private.Server/tfs/DefaultCollection" -TfsUserName "your_username" -TfsPassword "your_password"

.EXAMPLE
    # Using in Azure DevOps pipeline:
    # YAML pipeline:
    # - task: PowerShell@2
    #   inputs:
    #     filePath: '.\ConvertTo-Git.ps1'
    #     arguments: '-TfsPath "$/YourProject" -OutputPath "$(Build.ArtifactStagingDirectory)" -TfsCollection "https://Some.Private.Server/tfs/DefaultCollection" -TfsUserName "$(TfsUserName)" -TfsPassword "$(TfsPassword)"'
    #   displayName: 'Convert TFVC to Git'

.NOTES
    File Name      : ConvertTo-Git.ps1
    Author         : Amund N. Letrud (Tryll)
    Prerequisite   : 
    - PowerShell 5.1 or later
    - Visual Studio with Team Explorer (2019 or 2022)
    - Git command-line tools
    - Appropriate permissions in TFS/Azure DevOps
    
    Security:
    - Credentials are handled securely using SecureString objects
    - Passwords are cleared from memory after use
    - Compatible with Azure DevOps pipeline secret variables

.LINK
    https://github.com/Tryll/TeamFoundationServer
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$TfsPath,

    [Parameter(Mandatory=$true)]
    [string]$TfsCollection,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [string]$TfsUserName,
    
    [Parameter(Mandatory=$false)]
    [string]$TfsPassword
)

# Check if running in Pipeline
$isInPipeline = $env:TF_BUILD -eq "True"

# Check for password in environment variable if not provided as parameter
if ([string]::IsNullOrEmpty($TfsPassword) -and -not [string]::IsNullOrEmpty($env:TfsPassword)) {
    $TfsPassword = $env:TfsPassword
    Write-Host "Using TfsPassword from environment variable" -ForegroundColor Cyan
}

# Check if required .NET assemblies are available
$vsPath = @(
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer",
    "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer",
    "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer",
    "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
)

$tfAssemblyFound = $false

foreach ($path in $vsPath) {
    $tfAssemblyPath = Join-Path -Path $path -ChildPath "Microsoft.TeamFoundation.VersionControl.Client.dll"
    if (Test-Path $tfAssemblyPath) {
        Write-Host "Found TFS assembly at: $tfAssemblyPath" -ForegroundColor Green
        Add-Type -Path $tfAssemblyPath
        
        # Also load other required assemblies
        $clientAssemblyPath = Join-Path -Path $path -ChildPath "Microsoft.TeamFoundation.Client.dll"
        if (Test-Path $clientAssemblyPath) {
            Add-Type -Path $clientAssemblyPath
        }
        
        $tfAssemblyFound = $true
        break
    }
}

if (-not $tfAssemblyFound) {
    Write-Host "Error: TFS client assembly not found. Please make sure Visual Studio with Team Explorer is installed." -ForegroundColor Red
    exit 1
}

# Check if Git is installed
try {
    $gitVersion = git --version
    Write-Host "Git is available: $gitVersion" -ForegroundColor Green
} catch {
    Write-Host "Error: Git is not available. Please make sure Git is installed and in your PATH." -ForegroundColor Red
    exit 1
}

# Create output directory if it doesn't exist
if (!(Test-Path $OutputPath)) {
    Write-Host "Creating output directory: $OutputPath" -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Initialize Git repository
Write-Host "Initializing Git repository in $OutputPath..." -ForegroundColor Cyan
Set-Location $OutputPath

if (!(Test-Path ".git")) {
    git init
    
    # Create .gitattributes file for proper line endings
    @"
# Set default behavior to automatically normalize line endings
* text=auto

# Force batch scripts to always use CRLF line endings
*.{cmd,[cC][mM][dD]} text eol=crlf
*.{bat,[bB][aA][tT]} text eol=crlf

# Force bash scripts to always use LF line endings
*.sh text eol=lf

# Binary files
*.png binary
*.jpg binary
*.gif binary
*.ico binary
*.zip binary
*.pdf binary
*.xlsx binary
*.docx binary
*.pptx binary
"@ | Out-File -FilePath ".gitattributes" -Encoding utf8
    
    git add .gitattributes
    git commit -m "Initial commit with .gitattributes"
}

# Connect to TFS with appropriate authentication
Write-Host "Connecting to TFS at $TfsCollection..." -ForegroundColor Cyan
$startTime = Get-Date
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# Ignoring self signed
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

try {
    # Determine authentication method
    if (-not [string]::IsNullOrEmpty($TfsUserName)) {
        # Use username/password auth
        Write-Host "Using username/password authentication for $TfsUserName" -ForegroundColor Cyan
        
        $securePassword = $null
        
        # Check if password was provided as parameter or environment variable
        if (-not [string]::IsNullOrEmpty($TfsPassword)) {
            # Convert string password to SecureString
            $securePassword = ConvertTo-SecureString -String $TfsPassword -AsPlainText -Force
            # Immediately clear the plain text password from memory
            $TfsPassword = $null
        }
        else {
            # Request password securely if not provided
            $securePassword = Read-Host "Enter password for $TfsUserName" -AsSecureString
        }
        
        $credentials = New-Object System.Net.NetworkCredential($TfsUserName, $securePassword)
        $tfsCred = New-Object Microsoft.TeamFoundation.Client.TfsClientCredentials(
            [Microsoft.TeamFoundation.Client.BasicAuthCredential]::new($credentials)
        )
        $tfsServer = New-Object Microsoft.TeamFoundation.Client.TfsTeamProjectCollection(
            [Uri]$TfsCollection, 
            $tfsCred
        )
        
        # Clear the secure password from memory
        if ($securePassword -ne $null) {
            $securePassword.Dispose()
        }
    }
    else {
        # Fall back to default/integrated Windows authentication
        Write-Host "Using default Windows authentication" -ForegroundColor Cyan
        $tfsServer = New-Object Microsoft.TeamFoundation.Client.TfsTeamProjectCollection(
            [Uri]$TfsCollection
        )
    }
    
    # Connect to server
    $tfsServer.Authenticate()
    $vcs = $tfsServer.GetService([Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer])
    
    Write-Host "Connected successfully" -ForegroundColor Green
} catch {
    Write-Host "Error connecting to TFS: $_" -ForegroundColor Red
    exit 1
}

# Get all changesets for the specified path
Write-Host "Retrieving history for $TfsPath (this may take a while)..." -ForegroundColor Cyan
$history = $vcs.QueryHistory(
    $TfsPath,
    [Microsoft.TeamFoundation.VersionControl.Client.VersionSpec]::Latest,
    0,
    [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::Full,
    $null,
    $null,
    $null,
    [int]::MaxValue,    # Get all changesets
    $true, # Include details
    $false # Don't include download info to improve performance
)

# Sort changesets by date (oldest first)
$sortedHistory = $history | Sort-Object CreationDate

$totalChangesets = $sortedHistory.Count
Write-Host "Found $totalChangesets changesets - processing from oldest to newest" -ForegroundColor Green

# Initialize counters
$processedChangesets = 0
$processedFiles = 0

# Process each changeset
foreach ($cs in $sortedHistory) {
    $processedChangesets++
    $changesetId=$cs.ChangesetId
    $progressPercent = [math]::Round(($processedChangesets / $totalChangesets) * 100, 2)
    if ($isInPipeline) {
        Write-Host "##vso[task.setprogress value=$progressPercent;]Changeset $changesetId # $processedChangesets / $totalChangesets ($progressPercent%)"
    } else {
        Write-Progress -Activity "Replaying" -Status "Changeset $changesetId # $processedChangesets / $totalChangesets ($progressPercent%)" -PercentComplete $progressPercent
    }

    Write-Host "[TFVC-$changesetId] Processing by $($cs.OwnerDisplayName) from $($cs.CreationDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
    
    # Get detailed changeset info
    $changeset = $vcs.GetChangeset($cs.ChangesetId)
    $changeCount = $changeset.Changes.Count
    Write-Host "[TFVC-$changesetId] Contains $changeCount changes" -ForegroundColor Gray
   

    # Process each change in the changeset
    $changeCounter=0
    foreach ($change in $changeset.Changes) {
        $changeCounter++
        $changeItem = $change.Item
        $changesetId = [Int]::Parse($changeItem.ChangesetId)
        $changeType = [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]($change.ChangeType)
        $itemId= [Int]::Parse($changeItem.ItemId)
        $itemPath = $changeItem.ServerItem
        $relativePath = $itemPath.Substring($TfsPath.Length).TrimStart('/').Replace('/', '\')
        
        if ($changeItem.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Folder -or $changeItem.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Any) {

            if ($relativePath -ne "") {
                Write-Host "[TFVC-$changesetId] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray
                $d = mkdir $relativePath -ErrorAction SilentlyContinue
                continue
            }
            
        }

        #$change | ConvertTo-Json
     

        # Proces changes only if htey have a file path
        if (-not [String]::IsNullOrEmpty($relativePath)) {
         
            $localPath = Join-Path -Path $OutputPath -ChildPath $relativePath

            # Handle different change types
            switch ($change.ChangeType) {
                { $_ -band ([Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Add -bor 
                    [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Edit -bor
                    [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Rename)} {

                   
                    Write-Host "[TFVC-$changesetId] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray

                   
                    # Create directory structure
                    $localDir = Split-Path -Path $localPath -Parent
                    if (!(Test-Path $localDir)) {
                        New-Item -ItemType Directory -Path $localDir -Force | Out-Null
                    }
                    
                    # Download the file if it's not a directory
                    if ($change.Item.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::File) {
                        try {
                            $item = $vcs.GetItem($itemId, $changesetId)
                            $item.DownloadFile($localPath)
                            $processedFiles++
                        } catch {
                            Write-Host "[TFVC-$changesetId] [$changeCounter/$changeCount] Warning: Failed to download ${itemPath} [$changesetId/$itemId]: $_" -ForegroundColor Yellow
                        }
                    }
                    break
                }
                
                { $_ -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete } {

                    Write-Host "[TFVC-$changesetId] [$changeCounter/$changeCount] [Delete] $relativePath" -ForegroundColor Gray
                    
                    # Remove the file or directory
                    if (Test-Path $localPath) {
                        if (Test-Path $localPath -PathType Container) {
                            Remove-Item -Path $localPath -Recurse -Force
                        } else {
                            Remove-Item -Path $localPath -Force
                        }
                    }
                    break
                }
                default {
                    Write-Host "[TFVC-$changesetId] [$changeCounter/$changeCount] Unhandled change type: $($change.ChangeType) for $relativePath" -ForegroundColor Yellow
                    break
                }
            }
        }
    }
    # Commit changes to Git
    Write-Host "[TFVC-$changesetId] Committing changeset to Git" -ForegroundColor Gray
    
    # Stage all changes
    git add -A
    
    # Set environment variables for commit author and date
    $env:GIT_AUTHOR_NAME = $changeset.OwnerDisplayName
    $env:GIT_AUTHOR_EMAIL = "$($changeset.OwnerDisplayName.Replace(' ', '.'))"
    $env:GIT_AUTHOR_DATE = $changeset.CreationDate.ToString('yyyy-MM-dd HH:mm:ss K')
    $env:GIT_COMMITTER_NAME = $changeset.OwnerDisplayName
    $env:GIT_COMMITTER_EMAIL = "$($changeset.OwnerDisplayName.Replace(' ', '.'))"
    $env:GIT_COMMITTER_DATE = $changeset.CreationDate.ToString('yyyy-MM-dd HH:mm:ss K')
    
    # Prepare commit message
    $commitMessage = "[TFVC-$($changeset.ChangesetId)] $($changeset.Comment)"
    
    # Make the commit
    git commit -m $commitMessage --allow-empty
    
    # Clean up environment variables
    Remove-Item Env:\GIT_AUTHOR_NAME -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_AUTHOR_EMAIL -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_AUTHOR_DATE -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_COMMITTER_NAME -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_COMMITTER_EMAIL -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_COMMITTER_DATE -ErrorAction SilentlyContinue
    
    Write-Host "[TFVC-$changesetId] Completed" -ForegroundColor Green
}

# Clear the progress bar
if ($isInPipeline) {
    Write-Host "##vso[task.complete result=Succeeded;]DONE"
} else {
    Write-Progress -Activity "Replaying" -Completed
}

$endTime = Get-Date
$duration = $endTime - $startTime

Write-Host "`nConversion completed!" -ForegroundColor Green
Write-Host "Total changesets processed: $processedChangesets" -ForegroundColor Green
Write-Host "Total files processed: $processedFiles" -ForegroundColor Green
Write-Host "Total conversion time: $($duration.Hours) hours, $($duration.Minutes) minutes, $($duration.Seconds) seconds" -ForegroundColor Green
Write-Host "Git repository location: $OutputPath" -ForegroundColor Green
Write-Host "`nNext steps:" -ForegroundColor Cyan
Write-Host "1. Review the Git repository to ensure everything was migrated correctly" -ForegroundColor Cyan
Write-Host "2. Add a remote: git remote add origin <your-git-repo-url>" -ForegroundColor Cyan
Write-Host "3. Push to your Git repository: git push -u origin main" -ForegroundColor Cyan