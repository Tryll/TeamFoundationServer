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

    # The name of the TFS primary project branch and the default Git branch name
    [Parameter(Mandatory=$false)]
    [string]$PrimaryBranchName = "main",

    [Parameter(Mandatory=$false)]
    [int]$FromChangesetId = 0,

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
    [string]$AccessToken = $env:TfsAccessToken,
    
    [Parameter(Mandatory=$false)]
    [string]$LogFile = "$env:TEMP\merge-branch-$(Get-Date -Format 'yyyy-MM-dd-HHmmss').txt"

)


# Support functions
# ********************************
#region SupportFunctions

# find branch by path, longest to shortest
function Get-Branch  {
    param ($path)
    $currentPath = $path.Trim('/')
    while ($currentPath -ne "") {
        if ($branches.ContainsKey($currentPath)) {
            return $branches[$currentPath]
        }
        $currentPath = $currentPath.Substring(0, $currentPath.LastIndexOf('/'))
    }
    throw "Unknown branch for $path"
}

# create a new branch directly from container, input should never be a file.
function Add-Branch {
    param ($fromContainer)
    $fromContainer = $fromContainer.Trim('/')

    $source = get-branch($fromContainer)
    $sourceName = $source.Name
    $branchName = $fromContainer.Replace($projectPath,"").replace("/","-").Replace("$", "").Replace(".","-").Replace(" ","-").Trim('-')
    if (Test-Path $branchName) {
        Write-Host "Branch $branchName already exists" -ForegroundColor Gray
        return get-branch($newContainer)
    }
    
    Write-Host "Creating branch '$branchName' from '$sourceName'" -ForegroundColor Cyan
    $branches[$fromContainer] = @{
        Name = $branchName

        TfsPath = $fromContainer 

        # Ensure the top root is removed
        Rewrite = $fromContainer.Substring($projectPath.Length).Trim('/')
    }

    # Create the new branch folder from source branch
    push-location $sourceName
    git branch $branchName
    git worktree add "../$branchName" $branchName

    pop-location

    $branchCount++
    return $branches[$fromContainer]
}

function Get-ItemBranch {
    param ($path, $changesetId)


    $changeSet = new-object Microsoft.TeamFoundation.VersionControl.Client.ChangesetVersionSpec -argumentlist @($changesetId)

    do {
        $item = new-object Microsoft.TeamFoundation.VersionControl.Client.ItemIdentifier($path,  $changeSet )
        $branchObject = $vcs.QueryBranchObjects($item, [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::None)
        if ($branchObject -ne $null) {
            break
        }

        $path = $path.SubString(0, $path.LastIndexOf('/'))
    } while ($path.Length -gt 2)

    if ($path -eq  '$') {
        return $null
    }
    return $path
}

function Get-NormalizedHash {
    param ([string]$FilePath)
    
    # Normalize line endings and calculate hash in one go
    $fullPath = Resolve-Path -Path $FilePath | Select-Object -ExpandProperty Path
    $content = [System.IO.File]::ReadAllText($fullPath) -replace "`r`n", "`n"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    return [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace("-", "")
}

#endregion




# Start transcript if LogFile is provided
if ($LogFile) {
    Start-Transcript -Path $LogFile -Append -ErrorAction SilentlyContinue
    Write-Host "Logging to: $LogFile" -ForegroundColor Gray
}

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
Push-Location $OutputPath


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

$project = $vcs.GetTeamProject($TfsProject)
if ($project -eq $null) {
    Write-Host "Error: Project $TfsProject not found" -ForegroundColor Red
    exit 1
}
$projectPath=$project.ServerItem
$projectBranch = "main"
Write-Host "Found project $projectPath"


# Create the first main branch folder and initialize Git
$d=mkdir $projectBranch
push-location $projectBranch
git init -b $projectBranch
git commit -m "init" --allow-empty
pop-location

# Track all branches, with default branch first:
$branches = @{
    # The first and default branch and  way to catch all floating TFS folders
    "$projectPath" = @{
        Name = $PrimaryBranchName # The name of the branch in TFS and GIT
        TfsPath = "$projectPath" 
        # The tfspath will be renamed to this, prefixed with branchName for folder
        Rewrite = ""
    }

}
# track changes to branches, will git commit to each branch
$branchChanges = @{}
$branchCount = 0


$fromVersion = $null
if ($FromChangesetId -gt 0) {
    $fromVersion = new-object Microsoft.TeamFoundation.VersionControl.Client.ChangesetVersionSpec $FromChangesetId
}

# DOWNLOAD all TFS Project history
$history = $vcs.QueryHistory(
    $projectPath,
    [Microsoft.TeamFoundation.VersionControl.Client.VersionSpec]::Latest,
    0,
    [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::Full,
    $null,
    $fromVersion,
    $null,
    [int]::MaxValue,    # Get all changesets
    $false, # Don't Include details yet
    $false  # Don't include download info to improve performance
)

# Sort changesets by date (oldest first)
$sortedHistory = $history | Sort-Object CreationDate

$totalChangesets = $sortedHistory.Count
Write-Host "Found $totalChangesets changesets - processing from oldest to newest" -ForegroundColor Green

# Initialize counters
$processedChangesets = 0
$processedFiles = 0

$branchHashTracker = @{}

# Process each changeset
foreach ($cs in $sortedHistory) {
    $processedChangesets++
    $changesetId=$cs.ChangesetId
    $progressPercent = [math]::Round(($processedChangesets / $totalChangesets) * 100, 2)

    # Display / Progress updates
    if ($isInPipeline) {
        Write-Host "##vso[task.setprogress value=$progressPercent;]Changeset $changesetId # $processedChangesets / $totalChangesets ($progressPercent%)"
    } else {
        Write-Progress -Activity "Replaying" -Status "Changeset $changesetId # $processedChangesets / $totalChangesets ($progressPercent%)" -PercentComplete $progressPercent
    }

    Write-Host "[TFS-$changesetId] Processing by $($cs.OwnerDisplayName) from $($cs.CreationDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
    
    # Get detailed changeset info
    $changeset = $vcs.GetChangeset($cs.ChangesetId)
    $changes = $vcs.GetChangesForChangeset($cs.ChangesetId, $true,  [int]::MaxValue, $null, $null, $true)
    $changeCount = $changes.Count
    Write-Host "[TFS-$changesetId] Contains $changeCount changes" -ForegroundColor Gray
   

    # Process each change in the changeset
    $changeCounter=0
    $changesetId=0
        
    foreach ($change in $changes) {
        $changeCounter++
        $changeItem = $change.Item
        $changesetId = [Int]::Parse($changeItem.ChangesetId)
        $changeType = [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]($change.ChangeType)
        $itemType = [Microsoft.TeamFoundation.VersionControl.Client.ItemType]($changeItem.ItemType)
        $itemId= [Int]::Parse($changeItem.ItemId)
        $itemPath = $changeItem.ServerItem
        $itemContainer = $changeItem.ServerItem

        if ($changeItem.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::File) {
            $itemContainer = $itemContainer.Substring(0, $itemContainer.LastIndexOf('/'))

        }

        # Abort on mysterious change
        if ($change.MergeSources.Count -gt 1) {
            $change | convertto-json
            throw "Multiple merge sources is not supported"
        } 
              
        # Skip changes not in the specified path
        if ($itemPath.StartsWith($projectPath) -eq $false) {
            Write-Host "[TFS-$changesetId] [UNKNOWN] [$changeCounter/$changeCount] [$changeType] $itemPath - skipping, out of project" -ForegroundColor Yellow
            continue
        }
        if ($itemPath -eq $projectPath) {
            continue
        }  
    
        # Retrieve TFS Branch for item in changeset
        $tfsBranchPath=  Get-ItemBranch $itemPath $changesetId
        if ($tfsBranchPath -eq $null) {
           $tfsBranchPath = "$projectPath/$projectBranch"
        }

       
        # Check if we have a defined branch:
        $branch = get-branch($tfsBranchPath)

        # Check if we have a branch change:
        $gitPath =$branch.TfsPath
        if ($gitPath -eq $projectPath) {
            $gitPath+="/$projectBranch"
        }
        if ($branch -eq $null -or $gitPath -ne $tfsBranchPath) {
            $branch = Add-Branch($tfsBranchPath)
            $branchName=$branch.Name 
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $itemPath - Creating branch $branchName" -ForegroundColor Yellow
        }
        $branchName=$branch.Name
        
        # Find file relative path by branch name (folder) and item path replaced with branch local path.
        # This is the magic that will ensure we track the same files across branches.
        $relativePath = $itemPath.Replace($branch.TfsPath, $branch.Rewrite).TrimStart('/').Replace('/', '\')

        # Enter Branch:
        push-location $branchName
        $branchChanges[$branchName] = $true


        # Merging
        if ($change.MergeSources.Count -gt 0 -and ($change.ChangeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Merge)) {
            
            # Find container, branch base path
            $sourceBranchPath =  Get-ItemBranch $change.MergeSources[0].ServerItem $changesetId
            if ($sourceBranchPath -eq $null) {
                throw "Missing branch? $sourceBranchPath -eq $null"
            }
            $sourceBranch = get-branch($sourceBranchPath)
            $sourceBranchName = $sourceBranch.Name
             # Find actual checking hash
            $sourceChangesetId = $change.MergeSources[0].VersionTo
            $sourceChangesetIdFrom = $change.MergeSources[0].VersionFrom
            $sourcehash = $branchHashTracker["$sourceBranchName-$sourceChangesetId"]
            if ($sourceChangesetId -ne $sourceChangesetIdFrom) {
                Write-Host "Not Implemented: Source range merge $sourceChangesetIdFrom - $sourceChangesetId, using top range only for now." -ForegroundColor Yellow
            }
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Merging from [tfs-$sourceChangesetId][$sourceBranchName][$sourcehash]" -ForegroundColor Gray

      
            
            # Simple fix for Root
            $tfsPath = $sourceBranch.TfsPath
            if ($tfsPath -eq $projectPath) {
                $tfsPath+="/main"
            }
            # Check if we have a defined branch:
            if ($tfsPath -ne $sourceBranchPath) {
                throw "Missing branch? $tfsPath -ne $sourceBranchPath"
            }
           
            $sourceRelativePath = $change.MergeSources[0].ServerItem.Replace($sourceBranch.TfsPath, $sourceBranch.Rewrite).TrimStart('/').Replace('/', '\')

         

            # DELETE: Handle if this is just a delete, we will not link the deleted source file and the target file for deletion
            if ($change.ChangeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete) {
                git rm -f $relativePath
                pop-location #branch
                continue
            }

            # Takes current branch head, incase we need to revert a file
            $backupHead = git rev-parse HEAD  
            # Git checkout from hashes failes from time to time, forcing a recursive look for the file first to trigger cache update
            $commits = git rev-list --all -- $sourceRelativePath

            # CHECKOUT from HASH:
            git checkout $sourcehash -- $sourceRelativePath
            $checkoutSucceeded = $?
            if (-not $checkoutSucceeded) {
                Write-Host "Available commits for $sourceRelativePath"
                $commits | ForEach-Object { Write-Output $_ }
                throw "Failed git checkout"
            }

            # CHECKOUT RENAME: Source and Destination is not the same :
            if ($sourceRelativePath -ne $relativePath) {
                git mv -f $sourceRelativePath $relativePath
                # revert the original sourcerelativePath
                git checkout $backupHead -- $sourceRelativePath
            }


       
            if ($change.Item.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::File) {

                # EDIT DOWNLOAD file checked in:
                if ($change.ChangeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Edit) {
               
                    $changeItem.DownloadFile($relativePath)

                } else {

                # QUALITY CONTROL: 
                $checkedFileHash = Get-NormalizedHash -FilePath $relativePath
                $tmpFileName = "$env:TEMP\$relativePath"
                $changeItem.DownloadFile($tmpFileName)
                $tmpFileHash = Get-NormalizedHash -FilePath $tmpFileName

                if ($checkedFileHash -ne $tmpFileHash) {
                    Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Merging from [tfs-$sourceChangesetId][$sourceBranchName][$sourcehash] : $sourceRelativePath - File hash mismatch" -ForegroundColor Red
                    Write-Host $tmpFileName
                    throw "stop here"
                }
                }
            }
            
            # Next item!
            pop-location #branch
            continue
        }
             
        # Consistency Check: 
        if ($change.MergeSources.Count -gt 0 -and !(($change.ChangeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Rename) -bor ($change.ChangeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Branch))) {
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - MergeSource > 0" -ForegroundColor Yellow
            $change | convertto-json
            throw "stop here"
        }

       
        
        # Create Folder
        if ($changeItem.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Folder -or $changeItem.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Any) {

            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray

            # Add new directory
            New-Item -Path "$relativePath\.gitkeep" -ItemType File -Force | Out-Null
            git add "$relativePath\.gitkeep"

            # Next item!
            pop-location #branch
            continue
            
        }
    
        # Remove file
        if ($changeItem.ItemType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete) {

            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [Delete] $relativePath" -ForegroundColor Gray
            # Remove the file or directory
            git rm -f $relativePath

            # Next item!
            pop-location #branch
            continue
        }

  
        # Download the file if it's not a directory
        if ($change.Item.ItemType -ne [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::File) {
            $itemType=[Microsoft.TeamFoundation.VersionControl.Client.ItemType]($change.Item.ItemType)
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath is unhandled $itemType" -ForegroundColor Yellow
            throw("Unhandled")
        }

        # Handle rename where it exists
        if ($change.MergeSources.Count -gt 0 -and $changeItem.ItemType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Rename) {
            
            $sourcePath = $change.MergeSources[0].ServerItem.Replace($branch.TfsPath, $branch.Rewrite).TrimStart('/').Replace('/', '\')
            git mv -f $sourcePath $relativePath
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Renamed from $sourcePath" -ForegroundColor Gray
            # Next item!
            pop-location #branch

        }

        # Commit/PUT file:
        Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray

        if ($change.MergeSources.Count -gt 0) {
            # If we still have a source dump it 
            $change.MergeSources | convertto-json
        }

        try {
            # Create directory structure and empty file
            $target = New-Item -Path $relativePath -ItemType File -Force
            $changeItem.DownloadFile($target.FullName)
            git add $relativePath
            $processedFiles++
        } catch {
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath Error: Failed to download ${itemPath} [$changesetId/$itemId]: $_" -ForegroundColor Red
        }
  

        # Next item!
        pop-location #branch
    }


    # Set environment variables for commit author and date
    $env:GIT_AUTHOR_NAME = $changeset.OwnerDisplayName
    $env:GIT_AUTHOR_EMAIL = "$($changeset.OwnerDisplayName.Replace(' ', '.'))"
    $env:GIT_AUTHOR_DATE = $changeset.CreationDate.ToString('yyyy-MM-dd HH:mm:ss K')
    $env:GIT_COMMITTER_NAME = $changeset.OwnerDisplayName
    $env:GIT_COMMITTER_EMAIL = "$($changeset.OwnerDisplayName.Replace(' ', '.'))"
    $env:GIT_COMMITTER_DATE = $changeset.CreationDate.ToString('yyyy-MM-dd HH:mm:ss K')

    # Commit changes to Git
    foreach($branch in $branchChanges.Keys) {
        push-location $branch
       
        # Stage all changes
        git add -A
        
        # Prepare commit message
        $commitMessage = "$($changeset.Comment) [TFS-$($changeset.ChangesetId)]"
        
        # Make the commit
        git commit -m $commitMessage --allow-empty

        $branchHashTracker["$branch-$changesetId"] = git rev-parse HEAD
        $hash=$branchHashTracker["$branch-$changesetId"]
        Write-Host "[TFS-$changesetId] [$branch] [$hash] Comitted changes" -ForegroundColor Gray
        pop-location
    }

    # Clean up environment variables
    Remove-Item Env:\GIT_AUTHOR_NAME -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_AUTHOR_EMAIL -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_AUTHOR_DATE -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_COMMITTER_NAME -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_COMMITTER_EMAIL -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_COMMITTER_DATE -ErrorAction SilentlyContinue
    
    Write-Host "[TFS-$changesetId] Completed" -ForegroundColor Green
    # reset and loop
    $branchChanges = @{}
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
Write-Host "Total branches processed: $branchCount" -ForegroundColor Green
Write-Host "Total files processed: $processedFiles" -ForegroundColor Green
Write-Host "Total conversion time: $($duration.Hours) hours, $($duration.Minutes) minutes, $($duration.Seconds) seconds" -ForegroundColor Green
Write-Host "Git repository location: $OutputPath" -ForegroundColor Green
Write-Host "`nNext steps:" -ForegroundColor Cyan
Write-Host "1. Review the Git repository to ensure everything was migrated correctly" -ForegroundColor Cyan
Write-Host "2. Add a remote: git remote add origin <your-git-repo-url>" -ForegroundColor Cyan
Write-Host "3. Push to your Git repository: git push -u origin main" -ForegroundColor Cyan

pop-location


