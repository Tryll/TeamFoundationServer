<#
.SYNOPSIS
    Converts a TFVC repository to Git while preserving complete history and branch structure.

.DESCRIPTION
    ConvertTo-Git.ps1 extracts TFVC history and directly replays/applies it to a Git repository.
    This script processes all changesets chronologically through the entire branch hierarchy,
    maintaining original timestamps, authors, and comments. It creates a consistent project rooted directory tree 
    suitable for large projects with complex branch structures.

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

.PARAMETER PrimaryBranchName
    The name of the TFS primary project branch and the default Git branch name.
    Defaults to "main" if not specified.

.PARAMETER FromChangesetId
    Starting changeset ID for migration. If specified, only changes from this ID forward will be processed.
    Defaults to 0 (process all changesets).

.PARAMETER WithQualityControl
    Switch parameter to enable additional validation. When enabled, each file version is verified against
    the source TFVC repository to ensure integrity. This will slow down the process but ensures accuracy.

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

.PARAMETER LogFile
    Path to write script execution log. Defaults to a timestamped file in the temp directory.

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
    # Starting migration from a specific changeset:
    .\ConvertTo-Git.ps1 -TfsProject "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://dev.azure.com/organization" -FromChangesetId 1000 -UsePAT -AccessToken "your-personal-access-token"

.EXAMPLE
    # Using quality control for data verification:
    .\ConvertTo-Git.ps1 -TfsProject "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://dev.azure.com/organization" -WithQualityControl -UsePAT -AccessToken "your-personal-access-token"

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

    # Quality control effectively checks every iteration of a file, this will slow down the process, but ensure the files are correct.
    [Parameter(Mandatory=$false)]
    [switch]$WithQualityControl,

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
    [string]$LogFile = "$env:TEMP\convertto-git-$(Get-Date -Format 'yyyy-MM-dd-HHmmss').txt"

)



# Support functions
# ********************************
#region SupportFunctions

# find branch by path, longest to shortest
function Get-GitBranch  {
    param ($path)
    if ($path -eq $null) {
        throw ("Get-GitBranch : Path cannot be null")
    }

    $currentPath = $path.Trim('/')
    while ($currentPath -ne "") {
        if ($branches.ContainsKey($currentPath)) {
            return $branches[$currentPath]
        }
        $currentPath = $currentPath.Substring(0, $currentPath.LastIndexOf('/'))
    }

    throw "Get-GitBranch Unknown branch for $path"
}

# create a new branch directly from container, input should never be a file.
function Add-GitBranch {
    param ($fromContainer)
    $fromContainer = $fromContainer.Trim('/')

    $source = Get-GitBranch($fromContainer)
    $sourceName = $source.Name
    $branchName = $fromContainer.Replace($projectPath,"").replace("/","-").Replace("$", "").Replace(".","-").Replace(" ","-").Trim('-')
    if (Test-Path $branchName) {
        Write-Verbose "Add-GitBranch: Branch $branchName already exists"
        return Get-GitBranch($newContainer)
    }
    
    Write-Verbose "Add-GitBranch: Creating branch '$branchName' from '$sourceName'"
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
    $succeeded = $?
    if (-not $succeeded) {
        throw ("Add-GitBranch: Work tree creation failed, to long paths? ")
    }
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
  
    $fullPath = (get-item -Path $FilePath).FullName
    $content = [System.IO.File]::ReadAllText($fullPath)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    return [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace("-", "")
}

function Ensure-ItemDirectory {
    param($itemType, $relativePath)

    $itemFolder = "."
    if ($itemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Folder) {
        $itemFolder = $relativePath 
    } else {
        $itemFolder = (Split-Path -Path $relativePath -Parent)
    } 
    # Create directory if it does not exist
    if (-not (Test-Path -Path $itemFolder)) {
        New-Item -Path (Join-Path $itemFolder '.gitkeep') -ItemType File -Force | Out-Null
    }
   
    return $itemFolder
}

# Get-CommitFileName: Case-insensitive search for files in Git commits.
# Returns properly-cased filename from a specific commit.
# Usage: $filename = Get-CommitFileName -commit "a1b2c3d" -path "readme.md"
# Then: git checkout a1b2c3d -- $filename
# Parameters: commit (hash), path (filename to find), insensitive (default: true)
function Get-CommitFileName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$commit,
        
        [Parameter(Mandatory=$true)]
        [string]$path,
        
        [switch]$insensitive = $true
    )
    $path =$path.replace("\", "/")
    # Use git show instead of git log to target a specific commit
    $out = git ls-tree -r --name-only $commit 2>&1
    if ($out -is [System.Management.Automation.ErrorRecord]) {
        throw $out
    }

    # Files will come first in the reverse order before hitting empty lines/git comment
    foreach ($file in $out) {
      
        # Case-insensitive comparison
        if ($insensitive) {
            if ([string]::Equals($file, $path, [StringComparison]::OrdinalIgnoreCase)) {
                if ($file -ne $path) {
                    Write-Verbose "Get-CommitFileName: Found Git commited cased file name '$file' for '$path'"
                }
                return $file
            }
        } else {
            if ($file -eq $path) {
                return $file
            }
        }
        
    }

    # Not found - show the full commit info for debugging
    git ls-tree -r --name-only $commit
    $str="Get-CommitFileName: '$path' not found in commit $commit"
    throw($str)
}


function Get-SourceItem {
    param($change, $changesetId)

    if ($change.MergeSources -eq $null -or $change.MergeSources.Count -eq 0) {
        throw "Get-SourceItem: change item has no sources - $changesetId $($change.Item.ServerItem)"
    }
   
    # Find container, branch base path
    $Source = @{
        Path = $change.MergeSources[0].ServerItem
    }
    if ($Source.Path -eq $null) {
        $item=$change.MergeSources[0].ServerItem
        Write-Verbose "Get-SourceItem failed finding branch for $item from $changesetId"
        throw "Get-SourceItem: Missing branch for $item"
    }

    $Source.BranchPath = Get-ItemBranch $Source.Path $changesetId
    $Source.Branch = Get-GitBranch $Source.BranchPath
    $Source.BranchName = $Source.Branch.Name
    $Source.ChangesetId = $change.MergeSources[0].VersionTo
    $Source.ChangesetIdFrom = $change.MergeSources[0].VersionFrom
    $Source.Hash = $branchHashTracker["$($Source.BranchName)-$($Source.ChangesetId)"]
    Write-Verbose "Get-SourceItem: [$($Source.BranchPath)]:[$($Source.Branch.TfsPath)] is [$($Source.BranchName)] [$($Source.ChangesetId)] with local $($Source.Branch.Rewrite)"
  
    if ($Source.Hash -eq $null) {
        throw ("Source hash cannot be null")
    }

    if ($Source.ChangesetId -ne $Source.ChangesetIdFrom) {
        Write-Verbose "Get-SourceItem: Not Implemented: Source range merge $($Source.ChangesetId) - $($Source.ChangesetIdFrom), using top range only for now."
    }
    
    # Simple fix for Root, using global $projectPath
    if ($Source.Branch.TfsPath -eq $projectPath) {
        $Source.Branch.TfsPath+="/main"
    }
    $Source.RelativePath = $Source.Path.Replace($Source.Branch.TfsPath, $Source.Branch.Rewrite).TrimStart('/').Replace('/', '\')

    Write-Verbose "Get-SourceItem: [$($Source.BranchName)] [$($Source.ChangesetId)-$($Source.ChangesetIdFrom)] [$($Source.Hash)] $($Source.RelativePath)"

    # This may fail if a changeset has an add and an move in it of the same file, but it is a rare case.
    $Source.RelativePath  = Get-CommitFileName -commit  $Source.Hash -path $Source.RelativePath 

    return $Source
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

# Default Git settings
git config core.autocrlf false
git config core.longpaths true
# Old TFS checkins are case-insensitive, so we need to ignore case.
git config core.ignoreCase true

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

    Write-Host "[TFS-$changesetId] Processing Changeset $changesetId, $($cs.OwnerDisplayName) @ $($cs.CreationDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
    
    # Get detailed changeset info
    $changeset = $vcs.GetChangeset($cs.ChangesetId)
    $changes = $vcs.GetChangesForChangeset($cs.ChangesetId, $true,  [int]::MaxValue, $null, $null, $true)
    $changeCount = $changes.Count
    Write-Host "[TFS-$changesetId] Changeset with $changeCount changes" -ForegroundColor Gray
   

    # Process each change in the changeset
    $changeCounter=0
    $changesetId=0
    $relativePath =""
        
    foreach ($change in $changes) {

        $changeCounter++
        $changeItem = $change.Item
        $changesetId = [Int]::Parse($changeItem.ChangesetId)
        $changeType = [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]($change.ChangeType)
        $itemType = [Microsoft.TeamFoundation.VersionControl.Client.ItemType]($changeItem.ItemType)
        $itemId= [Int]::Parse($changeItem.ItemId)
        $itemPath = $changeItem.ServerItem
    

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
        $branch = Get-GitBranch($tfsBranchPath)

        # Check if we have a branch change:
        $gitPath =$branch.TfsPath
        if ($gitPath -eq $projectPath) {
            $gitPath+="/$projectBranch"
        }
        if ($branch -eq $null -or $gitPath -ne $tfsBranchPath) {
            $branch = Add-GitBranch($tfsBranchPath)
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

        try { #  try/finally for pop-location and  quality control


            # Merging/Branching and Rename (Rename essentially acts as a merge with source, since we are branching early)
            if (($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Merge -or
                 $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Branch -or
                 $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Rename)) {
                
                # The change item is a branch/merge with a source reference
                if ($change.MergeSources.Count -gt 0) {
                
                    # Lets ignore folders in merge/branch, as files are processed subsequently and git handles folders better/to good
                    if ($itemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Folder) {
                        Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Merging/Rename/Branch - ignoring container operations" -ForegroundColor Gray

                        # Next item!
                        continue
                    }

                    # DELETE: Handle if this is just a delete, we will not link the deleted source file and the target file for deletion
                    if ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete) {
                        Write-Verbose "Deleting $relativePath"
                        $out=git rm -f '$relativePath' 2>&1
                        if ($out -is [System.Management.Automation.ErrorRecord]) {
                            throw $out
                        }
                        continue
                    }


                    # Get source item
                    $source = Get-SourceItem $change $changesetId
                    $sourceBranchName = $source.BranchName
                    $sourceChangesetId = $source.ChangesetId
                    $sourcehash = $source.Hash
                    $sourceRelativePath = $source.RelativePath
                    
                    Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - from [tfs-$sourceChangesetId][$sourceBranchName][$sourcehash]" -ForegroundColor Gray
          

                    # Takes current branch head, incase we need to revert a file
                    $backupHead = $null
                    if (Test-Path -path $sourceRelativePath) {
                        # If file exists in target branch, we need to revert it back to original state
                        $backupHead = git rev-parse HEAD  
                    }
                    


                    # CHECKOUT from hash:
                    Write-Verbose "Checking out $sourceRelativePath from $sourcehash"
                    $out=git checkout -f $sourcehash -- '$sourceRelativePath' 2>&1
                    if ($out -is [System.Management.Automation.ErrorRecord]) {
                        throw $out
                    }


                    # CHECKOUT RENAME: Source and Destination is not the same : (GIT PROBLEMS:)
                    if ($sourceRelativePath -ne $relativePath) {

                        $dir=Ensure-ItemDirectory $itemType $relativePath
                        Write-Verbose "Renaming intermediate $sourceRelativePath to target $relativePath"
                        iex "git mv -fv '$sourceRelativePath' '$relativePath'"
                        if ($backupHead -ne $null) {
                            Write-Verbose "Reverting intermediate $sourceRelativePath"
                            # Revert the original sourcerelativePath
                            $out=git checkout -f $backupHead -- '$sourceRelativePath' 2>&1
                            if ($out -is [System.Management.Automation.ErrorRecord]) {
                                throw $out
                            }
                        }
                    }

                    # Let it continue to Edit!
                } else {

                    Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Without source, addding"
                    # Continue processing as normal file

                }
              
            }
     

            # Create Folder
            if ($itemType -band [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Folder) {

                Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray

                if ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete) {
                    Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Folder - ignoring container delete operations" -ForegroundColor Gray

                    # Next item!
                    continue
                }
                

                $d=Ensure-ItemDirectory $itemType $relativePath

                # Next item!
                continue
                
            }
        
            # Remove file
            if ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete -and -not ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::SourceRename)) {

                Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray
                # Remove the file or directory
                $out=git rm -f '$relativePath' 2>&1
                if ($out -is [System.Management.Automation.ErrorRecord]) {
                    throw $out
                }

                # Next item!
                continue
            }

            # SourceRename: Skip this item, as it will be handled by the next item
            if ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete -and $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::SourceRename) {
                Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - SourceRename (await next rename item)" -ForegroundColor Gray

                # Next item!
                continue
            }

            # Default Commit File action: Edit, Add, Branch without source and so on:
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray

            try {
                # Creates the target file and directory structure
                $target = New-Item -Path $relativePath -ItemType File -Force

                $changeItem.DownloadFile($target.FullName)

                $out=git add '$relativePath' 2>&1
                if ($out -is [System.Management.Automation.ErrorRecord]) {
                    throw $out
                }

                $processedFiles++

            } catch {
                Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath Error: Failed to download ${itemPath} [$changesetId/$itemId]: $_" -ForegroundColor Red
            }
        
        } catch {

            # On error disable qualitycontrol and exit
            $WithQualityControl = $false
            throw $_

        } finally {


            # QUALITY CONTROL: 
            if ($WithQualityControl -and $relativePath -ne "" -and ($itemType -ne [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Folder) -and
                 (-not ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete))) {

                Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] [$itemType] $relativePath - QC Processing"
                $checkedFileHash = Get-NormalizedHash -FilePath $relativePath
                $tmpFileName = "$env:TEMP\$relativePath"
                $changeItem.DownloadFile($tmpFileName)
                $tmpFileHash = Get-NormalizedHash -FilePath $tmpFileName

                if ($checkedFileHash -ne $tmpFileHash) {
                    Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Quality Control - File hash mismatch" -ForegroundColor Red
                    Write-Host $tmpFileName
                    throw "stop here"
                }

                Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - QC Pass" -ForegroundColor Gray
                remove-item -path $tmpFileName -force

            } 
              

            # Next item!
            pop-location #branch

        }
    }


    try {
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
            $branchHead = git rev-parse HEAD  

            $branchHashTracker["$branch-$changesetId"] = git rev-parse HEAD
            $hash=$branchHashTracker["$branch-$changesetId"]
            Write-Host "[TFS-$changesetId] [$branch] [$hash] Comitted" -ForegroundColor Gray
            pop-location

        }

    } finally {

        # Clean up environment variables
        Remove-Item Env:\GIT_AUTHOR_NAME -ErrorAction SilentlyContinue
        Remove-Item Env:\GIT_AUTHOR_EMAIL -ErrorAction SilentlyContinue
        Remove-Item Env:\GIT_AUTHOR_DATE -ErrorAction SilentlyContinue
        Remove-Item Env:\GIT_COMMITTER_NAME -ErrorAction SilentlyContinue
        Remove-Item Env:\GIT_COMMITTER_EMAIL -ErrorAction SilentlyContinue
        Remove-Item Env:\GIT_COMMITTER_DATE -ErrorAction SilentlyContinue
    }
   

    Write-Host "[TFS-$changesetId] Changeset Completed!" -ForegroundColor Green
    Write-Host ""
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


