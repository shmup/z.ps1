# configuration variables
$global:_Z_DATA = if ($env:_Z_DATA) { $env:_Z_DATA } else { "$env:USERPROFILE\.z" }
$global:_Z_CMD = if ($env:_Z_CMD) { $env:_Z_CMD } else { "z" }
$global:_Z_MAX_SCORE = if ($env:_Z_MAX_SCORE) { [int]$env:_Z_MAX_SCORE } else { 9000 }
$global:_Z_EXCLUDE_DIRS = if ($env:_Z_EXCLUDE_DIRS) { $env:_Z_EXCLUDE_DIRS -split ';' } else { @() }

function Get-ZDirs {
    if (Test-Path $global:_Z_DATA) {
        Get-Content $global:_Z_DATA | ForEach-Object {
            $parts = $_ -split '\|'
            if ($parts.Count -eq 3 -and (Test-Path $parts[0])) {
                $_
            }
        }
    }
}

function Add-ZEntry {
    param([string]$Path)
    
    # don't track home or root
    if ($Path -eq $env:USERPROFILE -or $Path -eq '\' -or $Path -eq '/') { return }
    
    # check excluded directories
    foreach ($exclude in $global:_Z_EXCLUDE_DIRS) {
        if ($Path -like "$exclude*") { return }
    }
    
    $tempFile = "$global:_Z_DATA.$([System.IO.Path]::GetRandomFileName())"
    $now = [DateTimeOffset]::Now.ToUnixTimeSeconds()
    $entries = @{}
    $totalCount = 0
    
    # read existing entries
    Get-ZDirs | ForEach-Object {
        $parts = $_ -split '\|'
        if ($parts.Count -eq 3) {
            $dir = $parts[0]
            $rank = [double]$parts[1]
            $time = $parts[2]
            
            if ($rank -ge 1) {
                if ($dir -eq $Path) {
                    $entries[$dir] = @{
                        rank = $rank + 1
                        time = $now
                    }
                } else {
                    $entries[$dir] = @{
                        rank = $rank
                        time = $time
                    }
                }
                $totalCount += $rank
            }
        }
    }
    
    # add new entry if not exists
    if (-not $entries.ContainsKey($Path)) {
        $entries[$Path] = @{
            rank = 1
            time = $now
        }
    }
    
    # write entries with aging if needed
    $output = @()
    foreach ($dir in $entries.Keys) {
        $rank = $entries[$dir].rank
        if ($totalCount -gt $global:_Z_MAX_SCORE) {
            $rank = $rank * 0.99
        }
        $output += "$dir|$rank|$($entries[$dir].time)"
    }
    
    $output | Out-File -FilePath $tempFile -Encoding utf8
    Move-Item -Path $tempFile -Destination $global:_Z_DATA -Force
}

function Get-Frecent {
    param(
        [double]$Rank,
        [long]$Time
    )
    $now = [DateTimeOffset]::Now.ToUnixTimeSeconds()
    $dx = $now - $Time
    return [int](10000 * $Rank * (3.75 / ((0.0001 * $dx + 1) + 0.25)))
}

function Invoke-Z {
    # get raw command line to avoid PowerShell parameter parsing
    $rawArgs = $args
    
    $list = $false
    $type = "frecent"
    $echo = $false
    $restrictPwd = $false
    $removeCurrentDir = $false
    $query = @()
    
    # parse arguments
    for ($i = 0; $i -lt $rawArgs.Count; $i++) {
        $arg = $rawArgs[$i]
        if ($arg -match '^-') {
            $opts = $arg.Substring(1)
            foreach ($opt in $opts.ToCharArray()) {
                switch ($opt) {
                    'l' { $list = $true }
                    'r' { $type = "rank" }
                    't' { $type = "recent" }
                    'e' { $echo = $true }
                    'c' { $restrictPwd = $true }
                    'x' { $removeCurrentDir = $true }
                    'h' { 
                        Write-Host "$global:_Z_CMD [-cehlrtx] args" -ForegroundColor Yellow
                        Write-Host "  -c  restrict matches to subdirs of current directory"
                        Write-Host "  -e  echo best match, don't cd"
                        Write-Host "  -h  show this help"
                        Write-Host "  -l  list matches instead of cd"
                        Write-Host "  -r  cd to highest ranked dir"
                        Write-Host "  -t  cd to most recently accessed dir"
                        Write-Host "  -x  remove current directory from datafile"
                        return
                    }
                }
            }
        } else {
            $query += $arg
        }
    }
    
    # remove current directory from datafile
    if ($removeCurrentDir) {
        if (Test-Path $global:_Z_DATA) {
            $pwd = (Get-Location).Path
            $lines = Get-Content $global:_Z_DATA | Where-Object { 
                -not ($_ -like "$pwd|*")
            }
            $lines | Out-File -FilePath $global:_Z_DATA -Encoding utf8
        }
        return
    }
    
    # build search pattern
    $searchPattern = if ($query.Count -gt 0) { 
        $pattern = ($query -join '.*')
        if ($restrictPwd) {
            $pwd = (Get-Location).Path
            "^$pwd.*$pattern"
        } else {
            $pattern
        }
    } else { 
        if ($restrictPwd) { 
            "^$(Get-Location).Path" 
        } else { 
            $null 
        }
    }
    
    if (-not $searchPattern -or $searchPattern -eq "^$(Get-Location).Path") {
        $list = $true
    }
    
    # no datafile yet
    if (-not (Test-Path $global:_Z_DATA)) { 
        Write-Host "No z data file found. cd around to build it up." -ForegroundColor Yellow
        return 
    }
    
    $now = [DateTimeOffset]::Now.ToUnixTimeSeconds()
    $matches = @{}
    $imatches = @{}
    $bestMatch = $null
    $iBestMatch = $null
    $hiRank = -999999999
    $iHiRank = -999999999
    
    # process entries
    Get-ZDirs | ForEach-Object {
        $parts = $_ -split '\|'
        if ($parts.Count -eq 3) {
            $dir = $parts[0]
            $rank = [double]$parts[1]
            $time = [long]$parts[2]
            
            # calculate score based on type
            $score = switch ($type) {
                "rank" { $rank }
                "recent" { $time - $now }
                default { Get-Frecent -Rank $rank -Time $time }
            }
            
            # check for matches
            if ($searchPattern) {
                if ($dir -match $searchPattern) {
                    $matches[$dir] = $score
                    if ($score -gt $hiRank) {
                        $bestMatch = $dir
                        $hiRank = $score
                    }
                } elseif ($dir -match $searchPattern.Replace('\*', '.*')) {
                    $imatches[$dir] = $score
                    if ($score -gt $iHiRank) {
                        $iBestMatch = $dir
                        $iHiRank = $score
                    }
                }
            } else {
                $matches[$dir] = $score
            }
        }
    }
    
    # output results
    if ($list) {
        # find common prefix
        $common = $null
        $shortest = $null
        foreach ($dir in $matches.Keys) {
            if ($matches[$dir] -and (-not $shortest -or $dir.Length -lt $shortest.Length)) {
                $shortest = $dir
            }
        }
        
        if ($shortest -and $shortest -ne '\' -and $shortest -ne '/') {
            $common = $shortest
            foreach ($dir in $matches.Keys) {
                if ($matches[$dir] -and -not $dir.StartsWith($shortest)) {
                    $common = $null
                    break
                }
            }
        }
        
        if ($common) {
            Write-Host "common:    $common" -ForegroundColor Cyan
        }
        
        # sort and display matches
        $sorted = if ($matches.Count -gt 0) { 
            $matches.GetEnumerator() | Sort-Object Value
        } else { 
            $imatches.GetEnumerator() | Sort-Object Value 
        }
        
        $sorted | ForEach-Object {
            Write-Host ("{0,-10} {1}" -f $_.Value, $_.Key)
        }
    } else {
        # navigate to best match
        $target = if ($bestMatch) { $bestMatch } elseif ($iBestMatch) { $iBestMatch } else { $null }
        
        if ($target) {
            if ($echo) {
                Write-Host $target
            } else {
                Set-Location $target
            }
        } else {
            Write-Host "No matches found" -ForegroundColor Yellow
        }
    }
}

# set up prompt hook to track directory changes
$global:_Z_LastLocation = $null

function Update-ZData {
    $currentLocation = (Get-Location).Path
    if ($currentLocation -ne $global:_Z_LastLocation) {
        $global:_Z_LastLocation = $currentLocation
        Add-ZEntry -Path $currentLocation
    }
}

# hook into prompt
$function:originalPrompt = $function:prompt
$function:prompt = {
    Update-ZData
    & $function:originalPrompt
}

# create alias
Set-Alias -Name $global:_Z_CMD -Value Invoke-Z -Scope Global

# tab completion
Register-ArgumentCompleter -CommandName $global:_Z_CMD -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    
    if (-not (Test-Path $global:_Z_DATA)) { return }
    
    $pattern = $wordToComplete -replace ' ', '.*'
    $dirs = @()
    
    Get-ZDirs | ForEach-Object {
        $parts = $_ -split '\|'
        if ($parts.Count -eq 3) {
            $dir = $parts[0]
            if ($dir -like "*$pattern*") {
                $dirs += $dir
            }
        }
    }
    
    $dirs | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_)
    }
}