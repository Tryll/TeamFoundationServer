param(
	$GitRepoUrl,
	$mainBranch="main"
)

if ([String]::IsNullOrEmpty($ENV:GIT_CONFIG_GLOBAL)) {
	$ENV:GIT_CONFIG_GLOBAL = join-path -Path $ENV:TEMP -ChildPath .gitconfig
}

if (-not (Test-Path $ENV:GIT_CONFIG_GLOBAL)) {
	New-Item -Path $ENV:GIT_CONFIG_GLOBAL -ItemType File -Force
}
git config --global  core.longpaths true



# Clone repo
git.exe clone -b $mainBranch $GitRepoUrl $mainBranch 
cd $mainBranch 

# Export all branches as folders
git.exe branch -r | ForEach-Object {

    $branch = $_.Trim() -replace '^origin/', ''
    if ($branch -ne 'HEAD' -and $branch -ne "$mainBranch" -and $branch -notmatch '->') {
	
	git.exe worktree add "../$branch" $branch
 	
    }
}
