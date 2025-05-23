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
    
    Performance:
    - For large repositories, this script may take several hours to run
    - Progress is displayed with percentage complete
    - Compatible with Azure DevOps pipeline tasks
    
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

    [Parameter(Mandatory=$false)]
    [string]$git = "git",


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
    param ($tfsPath)
    if ($tfsPath -eq $null) {
        throw ("Get-GitBranch : Path cannot be null")
    }

    $currentPath = $tfsPath.Trim('/')
    while ($currentPath -ne "" ) {
        if ($branches.ContainsKey($currentPath)) {
            return $branches[$currentPath]
        }
        # No more to check, default to main
        if ($currentPath.LastIndexOf('/') -lt 0) {
            break
        }
        $currentPath = $currentPath.Substring(0, $currentPath.LastIndexOf('/'))
    }

    # When we have processed all down to $/ the only implication is that we are looking for something not in the same project.
    Write-Verbose "Get-GitBranch: $tfsPath is from another project on TFS."
    throw ("Get-GitBranch: $tfsPath is from another project on TFS.")
}

# create a new branch directly from container, input should never be a file.
function Add-GitBranch {
    param ($fromContainer)
    $fromContainer = $fromContainer.Trim('/')

    # Find source
    $source = Get-GitBranch($fromContainer)
    $sourceName = $source.Name


    $branchName = $fromContainer.Replace($projectPath,"").replace("/","-").Replace("$", "").Replace(".","-").Replace(" ","-").Trim('-')
    if ($sourceName -eq  $branchName) {
        Write-Verbose "Add-GitBranch: Branch $branchName already exists"
        return $source
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
    # cygwin git unable to create workstrees. oh my.
    git branch $branchName | write-host
    
    git worktree add "../$branchName" $branchName | write-host
    $succeeded = $?
    if (-not $succeeded) {
        throw ("Add-GitBranch: Work tree creation failed, to long paths? ")
    }

    pop-location

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

    # Default to main
    if ($path -eq  '$') {
        Write-Verbose "Get-ItemBranch: Defaulting to main branch for $path at TFS-$changesetId"
        return "$projectPath/$projectBranch"
    }

    return $path
}

function Compare-Files {
    param (
        [Parameter(Mandatory=$true)]
        [string]$File1,
        
        [Parameter(Mandatory=$true)]
        [string]$File2
    )
    
    # Ensure the files exist
    if (-not (Test-Path -Path $File1)) {
        throw "File not found: $File1"
    }
    
    if (-not (Test-Path -Path $File2)) {
        throw "File not found: $File2"
    }
    
    # Get full paths to avoid any path-related issues
    $fullPath1 = (Resolve-Path $File1).Path
    $fullPath2 = (Resolve-Path $File2).Path
    
    # Use git diff with -w to ignore all whitespace differences (including BOMs)
    $result = & $git diff --no-index --exit-code -w $fullPath1 $fullPath2 2>&1
    
    # Git returns exit code 0 if files are identical, 1 if different
    # Return $true if files are the same (ignoring BOMs)
    return ($LASTEXITCODE -eq 0)
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

    if (-not $Source.BranchPath.StartsWith($projectPath)) {
        # We will support this from main thread, by always downloading - history will not be there though.
        Write-Verbose "Get-SourceItem: $($Source.BranchPath) is not from project $projectPath"
        throw("Get-SourceItem: Should not be here")
    }

    $Source.Branch = Get-GitBranch $Source.BranchPath
    $Source.BranchName = $Source.Branch.Name
    $Source.ChangesetId = $change.MergeSources[0].VersionTo
    $Source.ChangesetIdFrom = $change.MergeSources[0].VersionFrom
    $Source.Hash = $branchHashTracker["$($Source.BranchName)-$($Source.ChangesetId)"]
    # Simple fix for Root, using global $projectPath
    #if ($Source.Branch.TfsPath -eq $projectPath) {
    #    $Source.Branch.TfsPath+="/main"
    #}

    $Source.RelativePath = $Source.Path.Replace($Source.Branch.TfsPath, $Source.Branch.Rewrite).TrimStart('/').Replace('/', '\')

    Write-Verbose "Get-SourceItem: [$($Source.BranchPath)]:[$($Source.Branch.TfsPath)] is [$($Source.BranchName)] [$($Source.ChangesetIdFrom)-$($Source.ChangesetId)] with rewrite '$($Source.Branch.Rewrite)' for $($Source.RelativePath)"
  
    if ($Source.ChangesetId -ne $changesetId -and $Source.Hash -eq $null) {
        Write-Verbose "Get-SourceItem: Source Hash cannot be null"
    }

    if ($Source.ChangesetId -ne $Source.ChangesetIdFrom) {
        #Write-Verbose "Get-SourceItem: Not Implemented: Source range merge $($Source.ChangesetId) - $($Source.ChangesetIdFrom), using top range only for now."
        $lastFoundIn=0
        $lastFoundFile=""
        $gitLocalName = $Source.RelativePath.Replace("\","/")

        # Iterate from Top to Bottom, exit on first hit as this will be the newest change
        for($i= $Source.ChangesetId; $i -ge $Source.ChangesetIdFrom; $i--) {

            # Check if $i/"changesetid" is valid for this branch as a previous commit
            if ($branchHashTracker.ContainsKey("$($Source.BranchName)-$i")) {
                # Fetch hash from previous commit
                $tryHash = $branchHashTracker["$($Source.BranchName)-$i"]

                Write-Verbose "Get-SourceItem: Looking in $($Source.BranchName)-$i : $tryHash"
            
                # Check if file is changed and part of this commit
                $lastFoundFile = & $git show --name-only $tryHash 2>&1 | ? { $_ -ieq "$gitLocalName" }
                if ($lastFoundFile -ne $null ) {
                    $lastFoundIn=$i
                    # Break on first hit
                    break
                } else {
                    Write-Verbose "Did not find $gitLocalName in this:"
                    & $git show --name-only $tryHash 2>&1  | write-host
                    
                }
            }
        }
        if ($lastFoundIn -ne 0) {
            # using lastFoundFile to match case:
            Write-Verbose "Get-SourceItem: Scan found file ""$lastFoundFile"" last changed in TFS-$lastFoundIn"

            $Source.ChangesetId = $lastFoundIn
            # reverting back to windows format
            $Source.RelativePath = $lastFoundfile.Replace("/","\") 
            $Source.Hash = $branchHashTracker["$($Source.BranchName)-$($Source.ChangesetId)"]      
        }  else {
            Write-Verbose "Get-SourceItem: Scan failed to find $($Source.RelativePath) for changeset range $($Source.ChangesetId)-$($Source.ChangesetIdFrom)"
        }
    }
    

    Write-Verbose "Get-SourceItem: [$($Source.BranchName)] [$($Source.ChangesetId)-$($Source.ChangesetIdFrom)] [$($Source.Hash)] $($Source.RelativePath)"

 
    return $Source
}

<#
.SYNOPSIS
Sorts TFS change items to ensure Rename operations appear before their corresponding Add operations.

.DESCRIPTION
This function reorders an array of TFS change items by swapping positions of Rename and Add operations
for the same file paths, ensuring that Rename operations are processed before Add operations.
#>
function Sort-TfsChangeItems {
    param (
        [Parameter(Mandatory=$true)]
        [array]$changes
    )
    
    # Clone the array to avoid modifying the original
    $sorted = $changes.Clone()
    
    # Track Add files based on their path
    $addItems = @{}
    $idx = 0

    foreach ($change in $sorted) {

        # Track adds for files:
        if (($change.ChangeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Add) -and 
            ($change.Item.ItemType -band [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::File)) {
            
          #  Write-Verbose "Sort-TfsChangeItems tracking $($change.Item.ServerItem)"
            $addItems[$change.Item.ServerItem] = $idx++
            
        }   

        # Switch renames for files that match existing, so renames comes before adds
        if (($change.ChangeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Rename) -and 
            ($change.Item.ItemType -band [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::File) -and
            $change.MergeSources.Count -gt 0 -and 
            $addItems.ContainsKey($change.MergeSources[0].ServerItem)) {
            
        #    Write-Verbose "Sort-TfsChangeItems moving Rename before Add for $($change.MergeSources[0].ServerItem)"

            # Get the original Add change
            $origAddIdx = $addItems[$change.MergeSources[0].ServerItem]
            $origAddChange = $sorted[$origAddIdx]
            
            # Swap positions (put Rename before Add)
            $sorted[$origAddIdx] = $change
            $sorted[$idx] = $origAddChange
            
            # Update the index in the tracking dictionary
            $addItems[$change.MergeSources[0].ServerItem] = $idx
        } 
        
    }

    return $sorted
}


Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class PathAPI {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    [return: MarshalAs(UnmanagedType.U4)]
    public static extern int GetLongPathName(
        [MarshalAs(UnmanagedType.LPTStr)]
        string lpszShortPath,
        [MarshalAs(UnmanagedType.LPTStr)]
        StringBuilder lpszLongPath,
        [MarshalAs(UnmanagedType.U4)]
        int cchBuffer);
}
"@

function Get-RealCasedPath {
    param([string]$Path)
    
    $buffer = [System.Text.StringBuilder]::new(2048)
    $result = [PathAPI]::GetLongPathName($Path, $buffer,  $buffer.Capacity)
    
    if ($result -gt 0) {
        return $buffer.ToString(0, $result)
    }
    $buffer = $null
    return $Path  # fallback to original if API call fails
}

#endregion




# Start transcript if LogFile is provided
try { 
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




$gitVersion = & $git --version
Write-Host "Git version: $gitVersion" -ForegroundColor Green



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

# Create the first main branch folder and initialize Git
$d=mkdir $projectBranch
push-location $projectBranch
& $git init -b $projectBranch

& $git commit -m "init" --allow-empty

pop-location


$longPathsValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -ErrorAction SilentlyContinue
if ($null -eq $longPathsValue -or $longPathsValue.LongPathsEnabled -ne 1) {
    Write-Host "Warning: Long Paths not enabled for Windows!" -ForegroundColor Cyan
} else {
    Write-Host "Confirmed Long Paths is enabled for Windows"
}

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
$processedItems = 0

$gitGCCounter =0
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

    # Need to address "Add" vs "Rename" so that the order will be correct, Rename first the Add on the same file.
    # This will be handled as inplace replacement if discovered.
    Write-Verbose "Ensuring Rename < Add in changes"
    $ordered = Sort-TfsChangeItems -changes $changes


    foreach ($change in $ordered) {

        $changeCounter++
        $changeItem = $change.Item
        $changesetId = [Int]::Parse($changeItem.ChangesetId)
        $changeType = [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]($change.ChangeType)
        $itemType = [Microsoft.TeamFoundation.VersionControl.Client.ItemType]($changeItem.ItemType)
        $itemId= [Int]::Parse($changeItem.ItemId)
        $itemPath = $changeItem.ServerItem
        $processedItems++
        $forceAddNoSource = $false
        $fileDeleted = $false
        $qualityCheckNotApplicable = $false


        # Abort on mysterious change
        if ($change.MergeSources.Count -gt 1) {
            Write-Warning "[TFS-$changesetId]  [$changeCounter/$changeCount] [$changeType] $itemPath Has multiple merge sources"
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
        $path =$branch.TfsPath
        if ($path -eq $projectPath) {
            $path+="/$projectBranch"
        }
        if ($branch -eq $null -or $path -ne $tfsBranchPath) {
            $branch = Add-GitBranch($tfsBranchPath)
            $branchName=$branch.Name 
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $itemPath - Creating branch $branchName" -ForegroundColor Yellow
        }
        $branchName=$branch.Name
        
        # Find file relative path by branch name (folder) and item path replaced with branch local path.
        # This is the magic that will ensure we track the same files across branches.
        $relativePath = $itemPath.Replace($branch.TfsPath, $branch.Rewrite).TrimStart('/').Replace('/', '\')

        # Enter Branch:
        Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] [$itemType] $relativePath - Processing" -ForegroundColor Cyan
        push-location $branchName
        $branchChanges[$branchName] = $true

    

        try { #  try/finally for pop-location and  quality control

        


            # Retrieve files from a Source changeset, and prepare for other actions later
            # Merging/Branching and Rename (Rename essentially acts as a merge with source, since we are branching early)
            if (($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Merge -or
                 $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Branch -or
                 $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Rename -or
                 $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::SourceRename -or
                 $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Undelete -or
                 $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Rollback  )) {
                

                # Out of project route
                if ($change.MergeSources.Count -gt 0 -and -not $change.MergeSources[0].ServerItem.StartsWith($projectPath)) {
                    Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Is not from our project $projectPath!" -ForegroundColor Gray
                    $forceAddNoSource = $true
                }

                # TFS No Op in GIT
                if ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete -and 
                    $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::SourceRename) {

                    Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Delete+SourceRename is NO-OP" -ForegroundColor Gray
                    $qualityCheckNotApplicable = $true
                    #Next Item!
                    continue
                }


                # The change item is a branch/merge with a source reference
                if (-not $forceAddNoSource -and $change.MergeSources.Count -gt 0) {
                
                    # Lets ignore folders in merge/branch, as files are processed subsequently and git handles folders better/to good
                    if ($itemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Folder) {
                        Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Merging/Rename/Branch - ignoring container operations" -ForegroundColor Gray
                        $qualityCheckNotApplicable = $true
                        # Next item!
                        continue
                    }

                    # "Merge" operations on TFS without Edit or Branch is really nothing, and can be ignored - from the perspective of GIT.
                    if ($changeType -eq ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Merge)) {
                        Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Merging without Edit/Branch is a NO-OP in GIT" -ForegroundColor Gray
                        # There is nothing to check
                        $qualityCheckNotApplicable = $true

                        # Next item!
                        continue
                    }

                     # "Delete" + "Merge" + "SourceRename" => a file was renamed (and the source file "deleted") originally, there is nothing to track here as there is nothing to do.
                     if ($changeType -eq ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Merge -and 
                                          $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::SourceRename -and
                                          $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete)) {
                        Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Delete+Merge+SourceReName is a NO-OP in GIT" -ForegroundColor Gray
                        # There is nothing to check
                        $qualityCheckNotApplicable = $true

                        # Next item!
                        continue
                    }
                    



                    # Get source item
                    $source = Get-SourceItem $change $changesetId
                    $sourceBranchName = $source.BranchName
                    $sourceChangesetId = $source.ChangesetId
                    $sourcehash = $source.Hash
                    $sourceRelativePath = $source.RelativePath
                    
                    Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - from [tfs-$sourceChangesetId][$sourceBranchName][$sourcehash]" -ForegroundColor Gray
          

                    # Check if we are merging from another branch in the same changeset, this case would not allow checkout to function properly
                    if ($changesetId -eq  $sourceChangesetId -and $branchName -ne $sourceBranchName -and $sourcehash -eq $null) {
                        Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Reference is intra changeset, commiting early"
                 
                        pop-location # Exit current branch

                        try {
                            # Set environment variables for commit author and date
                            $env:GIT_AUTHOR_NAME = $changeset.OwnerDisplayName
                            $env:GIT_AUTHOR_EMAIL = "$($changeset.OwnerDisplayName.Replace(' ', '.'))"
                            $env:GIT_AUTHOR_DATE = $changeset.CreationDate.ToString('yyyy-MM-dd HH:mm:ss K')
                            $env:GIT_COMMITTER_NAME = $changeset.OwnerDisplayName
                            $env:GIT_COMMITTER_EMAIL = "$($changeset.OwnerDisplayName.Replace(' ', '.'))"
                            $env:GIT_COMMITTER_DATE = $changeset.CreationDate.ToString('yyyy-MM-dd HH:mm:ss K')
                    
                          
                            # Enter source branch for early commit
                            push-location $sourceBranchName
                            & $git status 2>&1 | Write-Host

                            & $git add -A 2>&1 | Write-Host
                            $commitMessage = "$($changeset.Comment) [TFS-$($changeset.ChangesetId)]"

                            $originalPreference = $ErrorActionPreference
                            $ErrorActionPreference = 'Continue'
                            & $git commit -m $commitMessage --allow-empty 2>&1 | Write-Host
                            $sourcehash = & $git rev-parse HEAD
                            $ErrorActionPreference = $originalPreference

                            $branchHashTracker["$sourceBranchName-$changesetId"] = $sourcehash
                            Write-Host "[TFS-$changesetId] [$sourceBranchName] [$sourcehash] Comitted" -ForegroundColor Gray

                            pop-location #sourceBranchName
                 
                            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - from [tfs-$sourceChangesetId][$sourceBranchName][$sourcehash] Commit updated!" -ForegroundColor Gray
          
                        

                        } finally {
                    
                            # Clean up environment variables
                            Remove-Item Env:\GIT_AUTHOR_NAME -ErrorAction SilentlyContinue
                            Remove-Item Env:\GIT_AUTHOR_EMAIL -ErrorAction SilentlyContinue
                            Remove-Item Env:\GIT_AUTHOR_DATE -ErrorAction SilentlyContinue
                            Remove-Item Env:\GIT_COMMITTER_NAME -ErrorAction SilentlyContinue
                            Remove-Item Env:\GIT_COMMITTER_EMAIL -ErrorAction SilentlyContinue
                            Remove-Item Env:\GIT_COMMITTER_DATE -ErrorAction SilentlyContinue
                        }

                        # Source branch dont need changeset finalization commit.
                        # IF this comes again, we'll overwrite branchHashTracker and be unable to refer to earlier commits from tfs.
                        # Expecting TFS to submitt its changes in sequence, so that this does not happen!
                        $branchChanges.Remove($sourceBranchName)
                        
                        # Reenter branch
                        push-location $branchName

                    }
                    


                    # Takes current branch head, incase we need to revert a file
                    $backupHead = $null
                    # Do not restore backup if move is in same changeset/branch/commit, the rename is rename
                    # Source has to exist
                    if ($sourcehash -ne $null -and (Test-Path -path $sourceRelativePath)) {
                        # If file exists in target branch, we need to revert it back to original state
                        $backupHead = & $git rev-parse HEAD  
                    }
                    
                    
                    # What do we do for branches created in same changeset that we want to copy from here!

                    # CHECKOUT from hash, it that exists - else file is local to branch:
                    if ($sourcehash -ne $null -and $changeItem.DeletionId -eq 0) {

                        # We need to flip this and manually find the appropiate case for git to be able to find the items.
                        $flipped=$sourceRelativePath.Replace("\","/") # Flip to linux path seps
                        $sourceRelativePath = & $git show --name-only $sourcehash 2>&1 | ? { $_ -ieq "$flipped" }
                        #$sourceRelativePath = $sourceRelativePath.Replace("/","\") # Flip path seps back

                        Write-Verbose "Checking out $sourceRelativePath from $sourcehash"

                        $originalPreference = $ErrorActionPreference
                        $ErrorActionPreference = 'Continue'
                        $out=& $git checkout -f $sourcehash -- "$sourceRelativePath" 2>&1
                        $ErrorActionPreference = $originalPreference

                        # We need to flip this back for this to work
                        $sourceRelativePath = $sourceRelativePath.Replace("/","\")
                      
                        if ($out -is [System.Management.Automation.ErrorRecord]) {

                            write-host ($out |convertto-json)
                            $status = & $git show --name-only $sourcehash 2>&1
                            
                            Write-verbose ($changeItem | convertto-json)
                            Write-Verbose "Something whent wrong with git checkout [$sourcehash] $sourceRelativePath"

                            throw ($out)

                          
                        }
                       
                    } else {
                        if ($changeItem.DeletionId -gt 0) {
                              
                                # Decision: Will not forward merge deleted items, by findit it and removing it.
                                # This could lead to a problem later when a file is request "undeleted", we'll have to look it up at that time.
                                # This approach keeps GIT history correct.
                                Write-Verbose "$sourceRelativePath is intended to be deleted"
                                $fileDeleted = $true
                                # Avoiding move processing
                                $sourceRelativePath = $relativePath
                                
                        } 
                    }
                    

                    # CHECKOUT RENAME: Source and Destination is not the same : (GIT PROBLEMS:)
                    if ($sourceRelativePath -ne $relativePath) {

                  
                        # Continue with normal rename
                        Write-Verbose "Renaming intermediate $sourceRelativePath to target $relativePath"

                        # Ensure folder structure exists, and remove the target file
                        $targetFile = new-item -path $relativePath -type file -force -erroraction SilentlyContinue 
                        remove-item -path $relativePath -force -erroraction SilentlyContinue | Out-Null

                        # Flip to linux style
                        $sourceRelativePath=$sourceRelativePath.Replace("\","/") # Flip to linux path seps
                        $relativePath = $relativePath.Replace("\","/")

                        Write-Verbose "Renaming intermediate native $sourceRelativePath to target $relativePath"
                        $out=& $git mv -f "$sourceRelativePath" "$relativePath"  2>&1 

                        Write-Host ($out | convertto-json)
                        
                        # Flip back to windows
                        $relativePath = $relativePath.Replace("/","\")
                        $sourceRelativePath = $sourceRelativePath.Replace("/","\") 
             
                        if ($backupHead -ne $null) {
                            Write-Verbose "Reverting intermediate $sourceRelativePath"
                            # Revert the original sourcerelativePath
                            $out=& $git checkout -f $backupHead -- "$sourceRelativePath" 2>&1
                            if ($out -is [System.Management.Automation.ErrorRecord]) {
                                Write-Verbose "git checkout failed $backupHead $sourceRelativePath"
                                throw $out
                            }
                        }
                    }


                    # Let it continue to Edit!
                } else {

                    Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Without source, addding"
                    # Continue processing as normal file
                    $forceAddNoSource = $true
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
                
                new-item -path $relativePath -type directory -force -erroraction SilentlyContinue | Out-Null

                # Next item!
                continue
                
            }
        
          # Add/Edit - Downloading:
            if ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Add -or 
                $changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Edit -or 
                $forceAddNoSource ) {
        
                # Default Commit File action: Edit, Add, Branch without source and so on:
                Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Downloading" -ForegroundColor Gray

                try {
                    # Creates the target file and directory structure
                    $target = new-item -path $relativePath -itemType File -force -erroraction silentlycontinue
                    remove-item -path $relativePath
                
                    $changeItem.DownloadFile($target.FullName)

                    if (-not (Test-Path -path $target.FullName)) {
                        Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Download failed, file not found"
                        throw "stop here"
                    }

                    # Flip to linux and notify git

                    # Looks like we may have to add the correct file path for the file here for git to not get index problems.
                    # Ie we need to resolve the actuall path
                    $realRealtivePath = Get-RealCasedPath -path $target.FullName
                     Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $realRelativePath - from $($target.FullName)"
                    $realRelativePath = $realRealtivePath.SubString($realRelativePath.Length - $relativePath.Length).Replace("\","/")
                    Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $realRelativePath - Real relative path, for git add"
                    
                    # Remove the file or directory  - Is this strictly required ?
                    $out=& $git add "$realRelativePath" 2>&1
    
                    if ($out -is [System.Management.Automation.ErrorRecord]) {
                        Write-Host ($out |convertto-json )
                        Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - File download add failed" 
                    } 

                    
                    $qualityCheckNotApplicable = $true
                    $fileDeleted = $false
                    

                } catch {
                    Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath Error: Failed to download ${itemPath} [$changesetId/$itemId] to $relativePath : $_" -ForegroundColor Red
                    throw("Failed to download $itemPath to $relativePath")
                }

                # This fails intermittently, maybe a add -A is better
                #$out=git add "$relativePath" 2>&1
                #if ($out -is [System.Management.Automation.ErrorRecord]) {
                #    Write-Verbose "Git add $relativePath failed, for $($target.FullName)"
                #    throw $out
                #}
            }



      
            # Remove file, as last step, but not on undelete/SourceRename
            if ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete -and
                 -not ($changeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::SourceRename)) {
                Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Deleting" -ForegroundColor Gray


                # Make the delete
                $originalPreference = $ErrorActionPreference
                $ErrorActionPreference = 'Continue'

                # Flip to linux
                $relativePath = $relativePath.Replace("\","/").Trim()
                # Remove the file or directory
                $out=& $git rm -f "$relativePath" 2>&1
                # Flip to windows
                $relativePath = $relativePath.Replace("/","\")


                $ErrorActionPreference = $originalPreference

                $fileDeleted = $true
                if ($out -is [System.Management.Automation.ErrorRecord]) {
                    Write-Host ($out |convertto-json )
                    Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - File allready deleted/missing. (TFS Supported)" 
                } 

            }
        
        } catch {

            # On error disable qualitycontrol and exit
            $WithQualityControl = $false
            throw $_

        } finally {


   
            # QUALITY CONTROL: 
            if ($WithQualityControl -and $relativePath -ne "" -and ($itemType -ne [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Folder)) {

                Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] [$itemType] [$fileDeleted] [$qualityCheckNotApplicable] $relativePath - QC Processing"

                # Check resulting file 
                $qcStatus ="Pass"
                if (-not $fileDeleted) {
                    if (-not $qualityCheckNotApplicable) {
                        $tmpFileName = "$env:TEMP\QCFile.tmp"

                        # Ensure previous file is not present
                        remove-item -path $tmpFileName -force -erroraction SilentlyContinue

                        $changeItem.DownloadFile($tmpFileName)
                        if (Test-Path -path $tmpFileName) {
                            
                            $originalFileLength = (Get-Item -Path $relativePath).Length
                            $downloadedFileLength = (Get-Item -Path $tmpFileName).Length
                                
                            if ($originalFileLength -gt 0 -and $downloadedFileLength -eq 0) {
                                # Based on current understanding, after review, this is a TFS inconsistency, and the file is not present after after tfs dump either. Ignoring
                                Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - QC - Downloaded 0 bytes from TFS, ignoring/corrupt TFS" 
                                $qcStatus = "Failed & Ignored"
                            } else {
                                if (-not (Compare-Files -file1 $relativePath -file2 $tmpFileName)) {
                                    Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - QC - File hash mismatch ($originalFileLength vs $downloadedFileLength)"
                                    Write-Verbose ($change | convertto-json)
                                                                        
                                    Write-Host $tmpFileName
                                    Write-Host (get-content $tmpFileName)
                                    
                                    throw "stop here"
                                }
                            }

                            # Cleanup
                            remove-item -path $tmpFileName -force -erroraction SilentlyContinue
                        } else {
                            Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - QC - Unable to download file from TFS, ignoring" 
                            $qcStatus = "Failed & Ignored"
                        }

                    } else {
                        # We have downloaded the file, so we dont need check it by downloading it again
                        $qcStatus = "N/A"
                    }
                } else {

                    # Check if deleted file is still present
                    if (Test-Path -path $relativePath) {
                        Write-Verbose "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - QC - File still exists"
                        throw "stop here"
                    }
                }
                
                Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - QC $qcStatus" -ForegroundColor Gray
            

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
            $o = & $git status 2>&1
            write-verbose ($o -join "`n")

            # Stage all changes
            $o = & $git add -A 2>&1
            write-verbose ($o -join "`n")

            # Prepare commit message
            $commitMessage = "$($changeset.Comment) [TFS-$($changeset.ChangesetId)]"
            
            # Make the commit
            $originalPreference = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'

                    
            & $git commit -m $commitMessage --allow-empty 2>&1 | Write-Host
            $hash = & $git rev-parse HEAD  
            $ErrorActionPreference  = $originalPreference


            $branchHashTracker["$branch-$changesetId"] =  $hash
            Write-Host "[TFS-$changesetId] [$branch] [$hash] Comitted" -ForegroundColor Gray
            pop-location

            $gitGCCounter++

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
    
   
    if ($gitGCCounter -gt 10) {
        $gitGCCounter = 0
        Write-Verbose "Performing git garbage collection, every 20'th commit"
        & $git gc --quiet
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
Write-Host "Total branches processed: $($branches.Keys.Count)" -ForegroundColor Green
Write-Host "Total files processed: $processedItems" -ForegroundColor Green
Write-Host "Total conversion time: $($duration.Hours) hours, $($duration.Minutes) minutes, $($duration.Seconds) seconds" -ForegroundColor Green
Write-Host "Git repository location: $OutputPath" -ForegroundColor Green
Write-Host "`nNext steps:" -ForegroundColor Cyan
Write-Host "1. Review the Git repository to ensure everything was migrated correctly" -ForegroundColor Cyan
Write-Host "2. Add a remote: git remote add origin <your-git-repo-url>" -ForegroundColor Cyan
Write-Host "3. Push to your Git repository: git push -u origin main" -ForegroundColor Cyan

pop-location

} finally {
    # Stop transcript if LogFile is provided
    if ($LogFile) {
        Stop-Transcript  -ErrorAction SilentlyContinue
    }
}

