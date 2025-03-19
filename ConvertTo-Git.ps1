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


$branches = @{
    
    # The first and default branch and  way to catch all floating TFS folders
    "$TfsProject" = @{
        Name = $PrimaryBranchName # The name of the branch in TFS and GIT
        TfsPath = "$TfsProject" 
        # The tfspath will be renamed to this, prefixed with branchName for folder
        Rewrite = ""
    }

}
# track changes to branches, will git commit to each branch
$branchChanges = @{}
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

# create a new branch from an existing branch, input should never be a file.
function Add-Branch {
    param ($fromContainer, $newContainer)

    $newContainer = $newContainer.Trim('/')
    $fromContainer = $fromContainer.Trim('/')

    $branchName = $newContainer.Replace($TfsProject,"").replace("/","-").Replace("$", "").Replace(".","-").Replace(" ","-").Trim('-')
    if (Test-Path $branchName) {
        Write-Host "Branch $branchName already exists" -ForegroundColor Gray
        return get-branch($newContainer)
    }


    $sourceBranch = Get-Branch($fromContainer)
    $sourceName = $sourceBranch.Name
    Write-Host "Branch '$sourceName' ($fromContainer) to '$branchName' ($newContainer)" -ForegroundColor Cyan
    $branches[$newContainer] = @{
        Name = $branchName

        TfsPath = $newContainer

        # Ensure the top root is removed
        Rewrite = $newContainer.Substring($TfsProject.Length).Trim('/')
    }

    # Create the new branch folder from source branch
    push-location $sourceName
    git branch $branchName
    git worktree add "../$branchName" $branchName

    pop-location
    return $branches[$newContainer]
}

# create a new branch directly from container, input should never be a file.
function Add-BranchDirect {
    param ($fromContainer)
    $fromContainer = $fromContainer.Trim('/')

    $source = get-branch($fromContainer)
    $sourceName = $source.Name
    $branchName = $fromContainer.Replace($TfsProject,"").replace("/","-").Replace("$", "").Replace(".","-").Replace(" ","-").Trim('-')
    if (Test-Path $branchName) {
        Write-Host "Branch $branchName already exists" -ForegroundColor Gray
        return get-branch($newContainer)
    }
    
    Write-Host "Direct creating branch '$branchName' from '$sourceName'" -ForegroundColor Cyan
    $branches[$fromContainer] = @{
        Name = $branchName

        TfsPath = $fromContainer 

        # Ensure the top root is removed
        Rewrite = $fromContainer.Substring($TfsProject.Length).Trim('/')
    }

    # Create the new branch folder from source branch
    push-location $sourceName
    git branch $branchName
    git worktree add "../$branchName" $branchName

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

    if ($path.Length -le 2) {
        return "$TfsProject/main"
    }
    # just need the path of the branch
    return $path
}

if (!(Test-Path ".git")) {
    # Create the first main branch folder and initialize Git
    $d=mkdir main
    push-location main
    git init -b main
    git commit -m "init" --allow-empty
    pop-location
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

$fromVersion = $null
if ($FromChangesetId -gt 0) {
    $fromVersion = new-object Microsoft.TeamFoundation.VersionControl.Client.ChangesetVersionSpec $FromChangesetId
}
$history = $vcs.QueryHistory(
    $TfsProject,
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
    $changeset = $vcs.GetChangeset($cs.ChangesetId)
    $changes = $vcs.GetChangesForChangeset($cs.ChangesetId, $true,  [int]::MaxValue, $null, $null, $true)
    $changeCount = $changes.Count
    Write-Host "[TFS-$changesetId] Contains $changeCount changes" -ForegroundColor Gray
   

    # Process each change in the changeset
    $changeCounter=0
    
        
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

        
        # Skip changes not in the specified path
        if ($itemPath.ToLower().StartsWith($TfsProject.ToLower()) -eq $false) {
            Write-Host "[TFS-$changesetId] [UNKNOWN] [$changeCounter/$changeCount] [$changeType] $itemPath - skipping, out of project" -ForegroundColor Yellow
            continue
        }
    
        # Retrieve TFS Branch for item in changeset
        $itemBranch=  Get-ItemBranch $itemPath $changesetId
        if ($itemBranch -eq $null) {
            throw "Missing branch? $itemBranch -eq $null"
        }
        # Check if we have a defined branch:
        $branch = get-branch($itemBranch)
        $branchName=$branch.Name 

        # Simple fix for Root
        $tfsPath =$branch.TfsPath
        if ($tfsPath -eq $TfsProject) {
            $tfsPath+="/main"
        }

        # If we dont, define it:
        if ($branch -eq $null -or $tfsPath -ne $itemBranch) {
            $branch = Add-BranchDirect($itemBranch)
            $branchName=$branch.Name 
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $itemPath - Creating branch $branchName" -ForegroundColor Yellow
            # Ensure the branch is tagged as having changes
            $branchChanges[$branchName] = $true
        }

        # Find file relative path by branch name (folder) and item path replaced with branch local path.
        # This is the magic that will ensure we track the same files across branches.
        $relativePath = $itemPath.Replace($branch.TfsPath, $branch.Rewrite).TrimStart('/').Replace('/', '\')
    
        # Not seen this yet, but adding to be sure we catch it and stop processing
        if ($change.MergeSources.Count -gt 1) {
            $change | convertto-json
            throw "Multiple merge sources is not supported"
        } 
     

        # Change Item: Merging
        if ($change.MergeSources.Count -gt 0 -and ($change.ChangeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Merge)) {
   

            # Find container, branch base path
            $sourceBranchPath =  Get-ItemBranch $change.MergeSources[0].ServerItem $changesetId
            if ($sourceBranchPath -eq $null) {
                throw "Missing branch? $sourceBranchPath -eq $null"
            }
            $sourceBranch = get-branch($sourceBranchPath)
            $sourceBranchName = $sourceBranch.Name
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - Merging from $sourceBranchName" -ForegroundColor Yellow

            
            # Simple fix for Root
            $tfsPath = $sourceBranch.TfsPath
            if ($tfsPath -eq $TfsProject) {
                $tfsPath+="/main"
            }
            # Check if we have a defined branch:
            if ($tfsPath -ne $sourceBranchPath) {
                throw "Missing branch? $tfsPath -ne $sourceBranchPath"
            }
           
            $sourceRelativePath = $change.MergeSources[0].ServerItem.Replace($sourceBranch.TfsPath, $sourceBranch.Rewrite).TrimStart('/').Replace('/', '\')

            # Tag changeset as having changes on this branch
            $branchChanges[$branchName] = $true
    
            
            # This should effectively link the source branch to the target branch on specific files
            push-location $branchName

            # Checkout the source file and move it to the target branch
            if ($sourceRelativePath -ne $relativePath) {
                git checkout $sourceBranch.Name -- $sourceRelativePath
                git mv -f $sourceRelativePath $relativePath
            } else {
                git checkout $sourceBranch.Name --  $relativePath
            }
            

            pop-location
            
            continue
        }
             
        # Debug out if not rename or branch, want to catch merge:
        if ($change.MergeSources.Count -gt 0 -and !(($change.ChangeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Rename) -bor ($change.ChangeType -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Branch))) {
            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath - MergeSource > 0" -ForegroundColor Yellow
            $change | convertto-json
            throw "stop here"
          }

        
        # Change Item: Create Folder
        if ($changeItem.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Folder -or $changeItem.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Any) {

            if ($relativePath -ne "") {
                Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray
                $branchChanges[$branch.Name] = $true
                $branchRelativePath = $branch.Name +'\' + $relativePath
                $d = mkdir $branchRelativePath -Force -ErrorAction SilentlyContinue
                "" > "$branchRelativePath\.gitkeep"
                continue
            }
            
        }
     

        # Change Item: Process File
        if (-not [String]::IsNullOrEmpty($relativePath)) {
            $branchRelativePath = $branch.Name +'\' + $relativePath
            # Handle different change types
            switch ($change.ChangeType) {
               
                # Remove file
                { $_ -band [Microsoft.TeamFoundation.VersionControl.Client.ChangeType]::Delete } {

                    Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [Delete] $relativePath" -ForegroundColor Gray
                    $branchChanges[$branch.Name] = $true
                    
                    # Remove the file or directory
                    if (Test-Path $branchRelativePath) {
                        if (Test-Path $branchRelativePath -PathType Container) {
                            Remove-Item -Path $branchRelativePath -Recurse -Force
                        } else {
                            Remove-Item -Path $branchRelativePath -Force
                        }
                    }
                    break
                }

                # No point handling Rename as it is staged with delete, just fetch it as normal.

                # Default handling - Download file
                default {
                 
                    Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath" -ForegroundColor Gray
                    $branchChanges[$branch.Name] = $true
                    $fullPath = Join-Path -path (pwd) -ChildPath $branchRelativePath
                   
                    # Create directory structure
                    $localDir = Split-Path -Path $branchRelativePath -Parent
                    if ($localDir -ne "" -and !(Test-Path $localDir)) {
                        New-Item -ItemType Directory -Path $localDir -Force | Out-Null
                    }
                    
                    # Download the file if it's not a directory
                    if ($change.Item.ItemType -eq [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::File) {
                        try {
                            $changeItem.DownloadFile($fullPath)
                            $processedFiles++
                        } catch {
                            Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath Error: Failed to download ${itemPath} [$changesetId/$itemId]: $_" -ForegroundColor Red
                        }
                    } else {
                        $itemType=[Microsoft.TeamFoundation.VersionControl.Client.ItemType]($change.Item.ItemType)
                        Write-Host "[TFS-$changesetId] [$branchName] [$changeCounter/$changeCount] [$changeType] $relativePath is unhandled $itemType" -ForegroundColor Yellow
                    }
                    break
                }
            }
        }
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
        Write-Host "[TFS-$changesetId] [$branch] Committing changeset to $branch" -ForegroundColor Gray
        # Stage all changes
        git add -A
        

        
        # Prepare commit message
        $commitMessage = "$($changeset.Comment) [TFS-$($changeset.ChangesetId)]"
        
        # Make the commit
        git commit -m $commitMessage --allow-empty

     

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
Write-Host "Total files processed: $processedFiles" -ForegroundColor Green
Write-Host "Total conversion time: $($duration.Hours) hours, $($duration.Minutes) minutes, $($duration.Seconds) seconds" -ForegroundColor Green
Write-Host "Git repository location: $OutputPath" -ForegroundColor Green
Write-Host "`nNext steps:" -ForegroundColor Cyan
Write-Host "1. Review the Git repository to ensure everything was migrated correctly" -ForegroundColor Cyan
Write-Host "2. Add a remote: git remote add origin <your-git-repo-url>" -ForegroundColor Cyan
Write-Host "3. Push to your Git repository: git push -u origin main" -ForegroundColor Cyan

pop-location
