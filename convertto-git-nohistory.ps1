<#
.SYNOPSIS
    Downloads the latest version of TFVC branches to a Git repository without history.

.DESCRIPTION
    ConvertTo-Git-NoHistory.ps1 downloads the current state of TFVC branches to a Git repository.
    Each TFS branch becomes a Git branch with a single commit containing the latest files.

.PARAMETER TfsProject
    The TFVC path to convert, in the format "$/ProjectName".

.PARAMETER TfsCollection
    The URL to your TFS/Azure DevOps collection.

.PARAMETER OutputPath
    The local folder where the Git repository will be created.

.PARAMETER PrimaryBranchName
    The name of the primary Git branch (defaults to "main").

.PARAMETER UseWindows, UseBasic, UsePAT, Credential, AccessToken, LogFile
    Authentication parameters (same as convertto-git.ps1)

.EXAMPLE
    .\convertto-git-nohistory.ps1 -TfsProject "$/ProjectName" -TfsCollection "https://dev.azure.com/organization" -OutputPath "C:\OutputFolder" -UsePAT -AccessToken "your-token"

.NOTES
    This script downloads only the latest version of files without preserving history.
#>
[CmdletBinding(DefaultParameterSetName="UseWindows")]
param(
    [Parameter(Mandatory=$true)]
    [string]$TfsProject,

    [Parameter(Mandatory=$true)]
    [string]$TfsCollection,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,

    [Parameter(Mandatory=$false)]
    [string]$PrimaryBranchName = "main",

    [Parameter(Mandatory=$false)]
    [string]$git = "git",

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
    [string]$LogFile = "$env:TEMP\convertto-git-nohistory-$(Get-Date -Format 'yyyy-MM-dd-HHmmss').txt"
)

# Support functions
function Get-RealCasedPath {
    param([string]$Path)
    
    $buffer = [System.Text.StringBuilder]::new(2048)
    $result = [PathAPI]::GetLongPathName($Path, $buffer,  $buffer.Capacity)
    
    if ($result -gt 0) {
        return $buffer.ToString(0, $result)
    }
    return $Path
}

# Create Git branch for TFS path
function Add-GitBranch {
    param ($fromContainer)
    $fromContainer = $fromContainer.Trim('/')

    # Generate branch name
    $branchName = $fromContainer.Replace($projectPath,"").replace("/","-").Replace("$", "").Replace(".","-").Replace(" ","-").Trim('-')
    if ($branchName -eq "") {
        $branchName = $PrimaryBranchName
    }
    
    Write-Verbose "Add-GitBranch: Creating branch '$branchName' for '$fromContainer'"

    $branches[$fromContainer] = @{
        Name = $branchName
        TfsPath = $fromContainer 
        Rewrite = $fromContainer.Substring($projectPath.Length).Trim('/')
    }

    # Create new branch worktree
    New-Item -ItemType Directory -Path $branchName -Force | Out-Null
    Push-Location $branchName
    & $git init -b $branchName
    & $git commit -m "Initial commit" --allow-empty
    Pop-Location

    return $branches[$fromContainer]
}

# Download latest files for a branch
function Download-TfsBranch {
    param($branch)
    
    $branchName = $branch.Name
    $tfsPath = $branch.TfsPath
    Write-Host "Downloading branch: $tfsPath to Git branch: $branchName" -ForegroundColor Cyan
    
    Push-Location $branchName
    
    try {
        # Get latest items
        $items = $vcs.GetItems($tfsPath, [Microsoft.TeamFoundation.VersionControl.Client.VersionSpec]::Latest, [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::Full)
        
        # Process folders
        foreach ($item in $items) {
            if ($item.ItemType -eq 'Folder') {
                $relativePath = $item.ServerItem.Replace($tfsPath, "").Trim('/').Replace('/', '\')
                if ($relativePath -ne "") {
                    New-Item -ItemType Directory -Path $relativePath -Force | Out-Null
                }
            }
        }
        
        # Process files
        foreach ($item in $items) {
            if ($item.ItemType -eq 'File') {
                $relativePath = $item.ServerItem.Replace($tfsPath, "").Trim('/').Replace('/', '\')
                $fullPath = Join-Path -Path (Get-Location) -ChildPath $relativePath
                
                # Ensure directory exists
                $directory = [System.IO.Path]::GetDirectoryName($fullPath)
                if (!(Test-Path -Path $directory)) {
                    New-Item -ItemType Directory -Path $directory -Force | Out-Null
                }
                
                # Download file
                $item.DownloadFile($fullPath)
            }
        }
        
        # Commit changes
        & $git add -A
        & $git commit -m "Initial import of $tfsPath from TFS without history"
    }
    finally {
        Pop-Location
    }
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

# Start transcript if LogFile is provided
try { 
    if ($LogFile) {
        Start-Transcript -Path $LogFile -Append -ErrorAction SilentlyContinue
        Write-Host "Logging to: $LogFile" -ForegroundColor Gray
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
            Add-Type -Path $tfAssemblyPath
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

    # Create output directory
    if (!(Test-Path $OutputPath)) {
        Write-Host "Creating output directory: $OutputPath" -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    # Initialize Git repository root
    Write-Host "Initializing target root $OutputPath..." -ForegroundColor Cyan
    Push-Location $OutputPath

    # Connect to TFS
    Write-Host "Connecting to TFS at $TfsCollection..." -ForegroundColor Cyan
    $startTime = Get-Date

    try {
        switch ($PSCmdlet.ParameterSetName) {
            "UseWindows" {
                if ($Credential -ne [System.Management.Automation.PSCredential]::Empty) {
                    $tfsCred = $Credential.GetNetworkCredential()
                } else {
                    $tfsCred = [System.Net.CredentialCache]::DefaultNetworkCredentials
                }
            }
            "UseBasic" { $tfsCred = $cred.GetNetworkCredential() }
            "UsePAT" { 
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
        $tfsServer.Authenticate()
        $vcs = $tfsServer.GetService([Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer])
        Write-Host "Connected successfully" -ForegroundColor Green
    } catch {
        Write-Host "Error connecting to TFS: $_" -ForegroundColor Red
        exit 1
    }

    # Get project details
    Write-Host "Retrieving project $TfsProject..." -ForegroundColor Cyan
    $project = $vcs.GetTeamProject($TfsProject)
    if ($project -eq $null) {
        Write-Host "Error: Project $TfsProject not found" -ForegroundColor Red
        exit 1
    }
    $projectPath = $project.ServerItem
    Write-Host "Found project $projectPath"

    # Initialize Git settings
    $env:GIT_CONFIG_GLOBAL = Join-Path -Path (Get-Location) -ChildPath ".gitconfig"
    & $git config --global user.email "tfs@git"
    & $git config --global user.name "TFS migration"
    & $git config --global core.autocrlf false
    & $git config --global core.longpaths true
    & $git config --global core.ignorecase true
    & $git config --global core.quotepath false
    & $git config --global --add safe.directory '*'

    # Track branches
    $branches = @{}

    # Create primary branch
    $primaryBranch = Add-GitBranch -fromContainer $projectPath
    $primaryBranch.Name = $PrimaryBranchName
    $branches[$projectPath] = $primaryBranch

    # Get all branch objects
    $branchObjects = $vcs.QueryBranchObjects($projectPath, [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::Full)

    # Create branches
    foreach ($branchObject in $branchObjects) {
        $branchPath = $branchObject.Properties.RootItem.Item
        if (-not $branches.ContainsKey($branchPath)) {
            $branch = Add-GitBranch -fromContainer $branchPath
            $branches[$branchPath] = $branch
        }
    }

    # Download latest files for each branch
    foreach ($branch in $branches.Values) {
        Download-TfsBranch -branch $branch
    }

    $endTime = Get-Date
    $duration = $endTime - $startTime

    Write-Host "`nConversion completed!" -ForegroundColor Green
    Write-Host "Total branches processed: $($branches.Count)" -ForegroundColor Green
    Write-Host "Total conversion time: $($duration.Hours) hours, $($duration.Minutes) minutes, $($duration.Seconds) seconds" -ForegroundColor Green
    Write-Host "Git repository location: $OutputPath" -ForegroundColor Green
    Write-Host "`nNext steps:" -ForegroundColor Cyan
    Write-Host "1. Review the Git repository" -ForegroundColor Cyan
    Write-Host "2. Add a remote: git remote add origin <your-git-repo-url>" -ForegroundColor Cyan
    Write-Host "3. Push branches: git push -u origin --all" -ForegroundColor Cyan

    Pop-Location
} finally {
    if ($LogFile) {
        Stop-Transcript -ErrorAction SilentlyContinue
    }
}
