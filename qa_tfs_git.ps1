param($tfsRepoDump, $gitRepoDump, [switch]$verbose)


# Git repo dump follows workspace format with main and branches as folders
# Git branches use "-" as folder spearator

$files = 0
$nonequal =0
$missingfolders =0
dir $gitRepoDump | % {
    
    $tfsRelativePath = $_.Name -replace "-", "/"
    $tfsPath = Join-Path $tfsRepoDump $tfsRelativePath/
    $gitPath = $_.FullName
    Write-Host "branch $($tfsRelativePath):"

    if (-not (Test-Path $tfsPath)) {
        $tfsRelativePath = $_.Name -replace "-", " "
        $tfsPath = Join-Path $tfsRepoDump $tfsRelativePath/
     
        Write-Host "Trying branch $($tfsRelativePath):"
    }

    if (Test-Path $tfsPath) {
        
        $tfsFiles = dir $tfsPath -Recurse | % { $_.FullName.Substring($tfsRepoDump.Length+1) }
        
        $tfsFiles | % {
            $relativeFilePath = $_

            $gitFilePath = Join-Path $gitPath $relativeFilePath
            $tfsFilePath = Join-Path $tfsRepoDump $relativeFilePath
    
            
            if (Test-Path $gitFilePath -PathType Leaf) {
                $files ++
                # Compare files
                $output = git diff --no-index -w --exit-code $tfsFilePath $gitFilePath 2>&1
                
                if ($LASTEXITCODE -ne 0) {
                    $nonequal++
                    Write-Host "! $relativeFilePath"
                } else {
                    if ($verbose) {
                        Write-Host "= $relativeFilePath"
                    }   
                }
            }
            else {
                if (Test-Path $gitFilePath -PathType Container) {
                        #Ignore folders
                }
                else {
                    # Ignore empty tfs folders
                    if ((Get-ChildItem $tfsFilePath -Recurse | Measure-Object).Count -eq 0) {
                           #Ignore empty tfs folders
                    } else {
                        $missingfolders++
                        Write-Host "- $relativeFilePath"
                        # Write-Host "File $relativeFilePath exists in TFS but not in Git"
                    }
                 
                }
            }
        }
    }
    else {
        Write-Host "TFS path $tfsPath does not exist"
    }
   
}

write-host "Compared $files files, and found $nonequal not equal and $missingfolders missing folders."