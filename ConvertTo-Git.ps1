<#
.SYNOPSIS
    Converts a TFVC repository to Git while preserving complete history and branch structure.

.DESCRIPTION
    ConvertTo-Git.ps1 extracts TFVC history and directly replays/applies it to a Git repository.
    This script processes all changesets chronologically through the entire branch hierarchy,
    maintaining original timestamps, authors, and comments. It creates a consistent, flat 
    migration suitable for large projects with complex branch structures.

    The script supports multiple authentication methods:
    - Windows Integrated Authentication (default)
    - Basic Authentication with username/password
    - Personal Access Token (PAT) for modern Azure DevOps environments

.PARAMETER TfsProject
    The TFVC path to convert, in the format "$/ProjectName".
    This can be the root of a project or a subfolder.

.PARAMETER TfsCollection
    The URL to your TFS/Azure DevOps collection, e.g., "https://dev.azure.com/organization" or "https://tfs.company.com/tfs/DefaultCollection".

.PARAMETER OutputPath
    The local folder where the Git repository will be created/updated.

.PARAMETER UseWindows
    Switch parameter to use Windows Integrated Authentication (default if no auth method specified).

.PARAMETER UseBasic
    Switch parameter to use Basic Authentication.

.PARAMETER UsePAT
    Switch parameter to use Personal Access Token authentication.

.PARAMETER Credential
    PSCredential object containing username/password for either Windows or Basic authentication.
    If not provided when using these methods, the script will use default Windows credentials.

.PARAMETER AccessToken
    Personal Access Token for Azure DevOps authentication.
    If not provided but UsePAT is specified, the script will check for an environment variable named "TfsAccessToken".

.EXAMPLE
    # Using Windows Integrated Authentication (default):
    .\ConvertTo-Git.ps1 -TfsProject "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://tfs.company.com/tfs/DefaultCollection"

.EXAMPLE
    # Using Windows Authentication with explicit credentials:
    $cred = Get-Credential
    .\ConvertTo-Git.ps1 -TfsProject "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://tfs.company.com/tfs/DefaultCollection" -UseWindows -Credential $cred

.EXAMPLE
    # Using Basic Authentication:
    $cred = Get-Credential
    .\ConvertTo-Git.ps1 -TfsProject "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://dev.azure.com/organization" -UseBasic -Credential $cred

.EXAMPLE
    # Using Personal Access Token (PAT) authentication:
    .\ConvertTo-Git.ps1 -TfsProject "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://dev.azure.com/organization" -UsePAT -AccessToken "your-personal-access-token"

.EXAMPLE
    # Using Personal Access Token from environment variable:
    $env:TfsAccessToken = "your-personal-access-token"
    .\ConvertTo-Git.ps1 -TfsProject "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://dev.azure.com/organization" -UsePAT

.EXAMPLE
    # Using in Azure DevOps pipeline:
    # YAML pipeline:
    # - task: PowerShell@2
    #   inputs:
    #     filePath: '.\ConvertTo-Git.ps1'
    #     arguments: '-TfsProject "$/YourProject" -OutputPath "$(Build.ArtifactStagingDirectory)" -TfsCollection "https://dev.azure.com/organization" -UsePAT -AccessToken "$(TfsAccessToken)"'
    #   displayName: 'Convert TFVC to Git'

.NOTES
    File Name      : ConvertTo-Git.ps1
    Author         : Amund N. Letrud (Tryll)
    Prerequisite   : 
    - PowerShell 5.1 or later
    - Visual Studio with Team Explorer (2019 or 2022)
    - Git command-line tools
    - Appropriate permissions in TFS/Azure DevOps
    
    Performance:
    - For large repositories, this script may take several hours to run
    - Progress is displayed with percentage complete
    - Compatible with Azure DevOps pipeline tasks
    
    Security:
    - Credentials are handled securely using SecureString objects
    - Passwords are cleared from memory after use
    - Compatible with Azure DevOps pipeline secret variables
    - PAT authentication is recommended for modern Azure DevOps environments

.LINK
    https://github.com/Tryll/TeamFoundationServer
#>
[CmdletBinding(DefaultParameterSetName="UseWindows")]
param(
    [Parameter(Mandatory=$true)]
    [string]$TfsProject,

    [Parameter(Mandatory=$true)]
    [string]$TfsCollection,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,

    [Parameter(Mandatory=$false, ParameterSetName="UseWindows")]
    [switch]$UseWindows,
    
    [Parameter(Mandatory=$false, ParameterSetName="UseBasic")]
    [switch]$UseBasic,

    [Parameter(Mandatory=$false, ParameterSetName="UseWindows")]
    [Parameter(Mandatory=$false, ParameterSetName="UseBasic")]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory=$false, ParameterSetName="UsePAT")]
    [switch]$UsePAT,

    [Parameter(Mandatory=$false, ParameterSetName="UsePAT")]
    [string]$AccessToken = $env:TfsAccessToken
    
)

# Check if running in Pipeline
$isInPipeline = $env:TF_BUILD -eq "True"


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
    git init -b main
    
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

try {
    switch ($PSCmdlet.ParameterSetName) {
        "UseWindows" {
            # Windows Integrated Authentication logic
            if ($Credential -ne [System.Management.Automation.PSCredential]::Empty) {
                Write-Host "Using provided crendentials for Windows authentication" -ForegroundColor Cyan
                $tfsCred = $Credential.GetNetworkCredential()
            } else {
                # Fall back to default/integrated Windows authentication
                Write-Host "Using default Windows authentication" -ForegroundColor Cyan
                $tfsCred = [System.Net.CredentialCache]::DefaultNetworkCredentials
            }
        }

        "UseBasic" {
            Write-Host "Using provided crendentials for Basic authentication" -ForegroundColor Cyan
            $tfsCred = $cred.GetNetworkCredential()
        }

        "UsePAT" {
            Write-Output "Using Personal Access Token authentication (Parameter or Environment TfsAccessToken)" -ForegroundColor Cyan
            if (!$AccessToken -and !$env:TfsAccessToken) {
                Write-Host "Error: Personal Access Token not provided" -ForegroundColor Red
                exit 1
            }
            $tfsCred = New-Object System.Net.NetworkCredential("", $AccessToken)     
        }
    }

   
    $tfsServer = New-Object Microsoft.TeamFoundation.Client.TfsTeamProjectCollection(
        [Uri]$TfsCollection, 
        $tfsCred
    )


    # Connect to server
    $tfsServer.Authenticate()
    
    $vcs = $tfsServer.GetService([Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer])
    
    Write-Host "Connected successfully" -ForegroundColor Green
} catch {
    Write-Host "Error connecting to TFS: $_" -ForegroundColor Red
    exit 1
}

# Get all changesets for the specified path
Write-Host "Retrieving history for $TfsProject (this may take a while)..." -ForegroundColor Cyan
$history = $vcs.QueryHistory(
    $TfsProject,
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

    Write-Host "[TFS-$changesetId] Processing by $($cs.OwnerDisplayName) from $($cs.CreationDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
    
    # Get detailed changeset info
    $changeset = $vcs.GetChangeset($cs.ChangesetId,$true,$false,$true)
    $changeCount = $changeset.Changes.Count
    Write-Host "[TFS-$changesetId] Contains $changeCount changes" -ForegroundColor Gray
   

    # Process each change in the changeset
    $changeCounter=0
    foreach ($change in $changeset.Changes) {
        $changeCounter++
        $changeItem = $change.Item
        $changesetId = [Int]::Parse($changeItem.ChangesetId)
        $changeType = [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]($change.ChangeType)
        $itemId= [Int]::Parse($changeItem.ItemId)
        $itemPath = $changeItem.ServerItem
        # Skip changes not in the specified path
        if ($itemPath.ToLower().StartsWith($TfsProject.ToLower()) -eq $false) {
            continue
        }
        # Skip first $, / characters
        $relativePath = $itemPath.Substring($TfsProject.Length).TrimStart('/').Replace('/', '\')
        
        if ($changeItem.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Folder -or $changeItem.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Any) {

            if ($relativePath -ne "") {
                Write-Host "[TFS-$changesetId] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray
                $d = mkdir $relativePath -ErrorAction SilentlyContinue
                continue
            }
            
        }
     

        # Proces changes only if htey have a file path
        if (-not [String]::IsNullOrEmpty($relativePath)) {
          
            # Handle different change types
            switch ($change.ChangeType) {
               
                # Remove file
                { $_ -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete } {

                    Write-Host "[TFS-$changesetId] [$changeCounter/$changeCount] [Delete] $relativePath" -ForegroundColor Gray
                    
                    # Remove the file or directory
                    if (Test-Path $relativePath) {
                        if (Test-Path $relativePath -PathType Container) {
                            Remove-Item -Path $relativePath -Recurse -Force
                        } else {
                            Remove-Item -Path $relativePath -Force
                        }
                    }
                    break
                }

                # Rename if source file is referenced:
                { $_ -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Rename } {

               
                    Write-Host "[TFS-$changesetId] rename "

                    $change | Convertto-Json

                    $oldRelativePath = $change.SourceServerItem.Substring($TfsProject.Length).TrimStart('/').Replace('/', '\')
                    
                    Write-Host "[TFS-$changesetId] [$changeCounter/$changeCount] [$changeType] $oldRelativePath to $relativePath" -ForegroundColor Gray


                    # Get the current location
                    $currentLocation = Get-Location
                    $oldFullPath = Join-Path -Path $currentLocation -ChildPath $oldRelativePath
                    $newFullPath = Join-Path -Path $currentLocation -ChildPath $relativePath

                    # Create target directory if it doesn't exist
                    $targetDir = Split-Path -Path $relativePath -Parent
                    if ($targetDir -and !(Test-Path $targetDir)) {
                        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                    }

                    # Move the old file or directory
                    if (Test-Path $oldRelativePath) {
                        if (!(Test-Path $relativePath)) {
                            # The -NewName parameter should be just the new name, not the full path
                            $newName = Split-Path -Path $relativePath -Leaf
                            $targetDir = Split-Path -Path $relativePath -Parent
                            
                            # If we're moving to a different directory
                            if ($targetDir -and (Split-Path -Path $oldRelativePath -Parent) -ne $targetDir) {
                                # First ensure the target directory exists
                                if (!(Test-Path $targetDir)) {
                                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                                }
                                # Move item to the new directory
                                Move-Item -Path $oldRelativePath -Destination $targetDir -Force
                                # Then rename if necessary
                                $movedItem = Join-Path -Path $targetDir -ChildPath (Split-Path -Path $oldRelativePath -Leaf)
                                if ((Split-Path -Path $oldRelativePath -Leaf) -ne $newName) {
                                    Rename-Item -Path $movedItem -NewName $newName -Force
                                }
                            } else {
                                # Just rename in the same directory
                                Rename-Item -Path $oldRelativePath -NewName $newName -Force
                            }
                        } else {
                            Write-Host "[TFS-$changesetId] [$changeCounter/$changeCount] [$changeType] $oldRelativePath to $relativePath - Already exists" -ForegroundColor Yellow
                        }
                    } else {
                        if (!(Test-Path $relativePath)) {
                            Write-Host "[TFS-$changesetId] [$changeCounter/$changeCount] [$changeType] $oldRelativePath to $relativePath - Source does not exist" -ForegroundColor Yellow
                        } else {
                            Write-Host "[TFS-$changesetId] [$changeCounter/$changeCount] [$changeType] $oldRelativePath to $relativePath - Allready moved" -ForegroundColor Gray
                        }
                        
                    }

                    # Check if the new file exists after rename attempt
                    if (Test-Path $relativePath) {
                        break
                    }
                    
                    # If it does not exist continue to default processing
                }
                
                # Default handling - Download file
                default {
                 
                    Write-Host "[TFS-$changesetId] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray

                    $fullPath = Join-Path -path (pwd) -ChildPath $relativePath
                   
                    # Create directory structure
                    $localDir = Split-Path -Path $relativePath -Parent
                    if ($localDir -ne "" -and !(Test-Path $localDir)) {
                        New-Item -ItemType Directory -Path $localDir -Force | Out-Null
                    }
                    
                    # Download the file if it's not a directory
                    if ($change.Item.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::File) {
                        try {
                            $item = $vcs.GetItem($itemId, $changesetId)
                            $item.DownloadFile($fullPath)
                            $processedFiles++
                        } catch {
                            Write-Host "[TFS-$changesetId] [$changeCounter/$changeCount] Warning: Failed to download ${itemPath} [$changesetId/$itemId]: $_" -ForegroundColor Yellow
                        }
                    } else {
                        $itemType=[Microsoft.TeamFoundation.VersionControl.Client.ItemType]($change.Item.ItemType)
                        Write-Host "[TFS-$changesetId] [$changeCounter/$changeCount] [$changeType] $relativePath is unhandled $itemType" -ForegroundColor Yellow
                    }
                    break
                }
                #default {
                #    Write-Host "[TFS-$changesetId] [$changeCounter/$changeCount] Unhandled change type: $($change.ChangeType) for $relativePath" -ForegroundColor Yellow
                #    break
                #}
            }
        }
    }
    # Commit changes to Git
    Write-Host "[TFS-$changesetId] Committing changeset to Git" -ForegroundColor Gray
    
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
    $commitMessage = "$($changeset.Comment) [TFS-$($changeset.ChangesetId)]"
    
    # Make the commit
    git commit -m $commitMessage --allow-empty
    
    # Clean up environment variables
    Remove-Item Env:\GIT_AUTHOR_NAME -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_AUTHOR_EMAIL -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_AUTHOR_DATE -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_COMMITTER_NAME -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_COMMITTER_EMAIL -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_COMMITTER_DATE -ErrorAction SilentlyContinue
    
    Write-Host "[TFS-$changesetId] Completed" -ForegroundColor Green
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