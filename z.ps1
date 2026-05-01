# Copyright (c) 2025. Licensed under the WTFPL license, Version 2
# Based on rupa/z - https://github.com/rupa/z

# maintains a jump-list of the directories you actually use
#
# INSTALL:
#     * put something like this in your PowerShell profile:
#         . C:\path\to\z.ps1
#     * cd around for a while to build up the db
#     * optionally:
#         set $_Z_CMD to change the command (default z).
#         set $_Z_DATA to change the datafile (default ~\.z).
#         set $_Z_MAX_SCORE lower to age entries out faster (default 9000).
#         set $_Z_EXCLUDE_DIRS to an array of directories to exclude.
#
# USE:
#     * z foo     # cd to most frecent dir matching foo
#     * z foo bar # cd to most frecent dir matching foo and bar
#     * z -r foo  # cd to highest ranked dir matching foo
#     * z -t foo  # cd to most recently accessed dir matching foo
#     * z -l foo  # list matches instead of cd
#     * z -e foo  # echo the best match, don't cd
#     * z -c foo  # restrict matches to subdirs of current directory
#     * z -x      # remove the current directory from the datafile
#     * z -h      # show a brief help message

# configuration variables
$global:_Z_DATA = if ($env:_Z_DATA) { $env:_Z_DATA } else { "$env:USERPROFILE\.z" }
$global:_Z_CMD = if ($env:_Z_CMD) { $env:_Z_CMD } else { "z" }
$global:_Z_MAX_SCORE = if ($env:_Z_MAX_SCORE) { [int]$env:_Z_MAX_SCORE } else { 9000 }
$global:_Z_EXCLUDE_DIRS = if ($env:_Z_EXCLUDE_DIRS) { $env:_Z_EXCLUDE_DIRS -split ';' } else { @() }

if (Test-Path -LiteralPath $global:_Z_DATA -PathType Container) {
    Write-Host "ERROR: z.ps1's datafile ($global:_Z_DATA) is a directory." -ForegroundColor Yellow
}

function Get-ZEntries {
    if (-not (Test-Path -LiteralPath $global:_Z_DATA -PathType Leaf)) {
        return
    }

    Get-Content -LiteralPath $global:_Z_DATA | ForEach-Object {
        $parts = $_ -split '\|', 3
        $rank = 0.0
        $time = 0L
        if (
            $parts.Count -eq 3 -and
            [double]::TryParse($parts[1], [ref]$rank) -and
            [long]::TryParse($parts[2], [ref]$time) -and
            (Test-Path -LiteralPath $parts[0] -PathType Container)
        ) {
            [pscustomobject]@{
                Path = $parts[0]
                Rank = $rank
                Time = $time
            }
        }
    }
}

function Remove-ZEntry {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $global:_Z_DATA -PathType Leaf)) {
        return
    }

    $lines = Get-Content -LiteralPath $global:_Z_DATA | Where-Object {
        $parts = $_ -split '\|', 2
        $parts.Count -gt 0 -and $parts[0] -ne $Path
    }

    $lines | Out-File -LiteralPath $global:_Z_DATA -Encoding utf8
}

function Test-ZRegexMatch {
    param(
        [string]$Path,
        [string]$Pattern,
        [switch]$IgnoreCase
    )

    $options = [System.Text.RegularExpressions.RegexOptions]::None
    if ($IgnoreCase) {
        $options = $options -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    }

    return [System.Text.RegularExpressions.Regex]::IsMatch($Path, $Pattern, $options)
}

function Get-ZCommonPrefix {
    param([hashtable]$Matches)

    $shortest = $null
    foreach ($path in $Matches.Keys) {
        if ($Matches[$path] -and (-not $shortest -or $path.Length -lt $shortest.Length)) {
            $shortest = $path
        }
    }

    if (-not $shortest -or $shortest -eq '\' -or $shortest -eq '/') {
        return $null
    }

    foreach ($path in $Matches.Keys) {
        if ($Matches[$path] -and -not $path.StartsWith($shortest, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $null
        }
    }

    return $shortest
}

function Find-ZMatches {
    param(
        [string[]]$Query,
        [string]$Type,
        [switch]$RestrictPwd
    )

    $queryText = $Query -join ' '
    $pwdPrefix = $null
    if ($RestrictPwd) {
        $pwdPrefix = '^{0}' -f [System.Text.RegularExpressions.Regex]::Escape((Get-Location).Path)
        $queryText = if ($queryText) { "$pwdPrefix $queryText" } else { "$pwdPrefix " }
    }

    $pattern = if ($queryText) { $queryText -replace ' ', '.*' } else { $null }
    $now = [DateTimeOffset]::Now.ToUnixTimeSeconds()
    $matches = @{}
    $imatches = @{}
    $bestMatch = $null
    $iBestMatch = $null
    $hiRank = -9999999999
    $iHiRank = -9999999999

    foreach ($entry in Get-ZEntries) {
        $score = switch ($Type) {
            'rank' { $entry.Rank }
            'recent' { $entry.Time - $now }
            default { Get-Frecent -Rank $entry.Rank -Time $entry.Time }
        }

        if (-not $pattern) {
            $matches[$entry.Path] = $score
            continue
        }

        if (Test-ZRegexMatch -Path $entry.Path -Pattern $pattern) {
            $matches[$entry.Path] = $score
            if ($score -gt $hiRank) {
                $bestMatch = $entry.Path
                $hiRank = $score
            }
        } elseif (Test-ZRegexMatch -Path $entry.Path -Pattern $pattern -IgnoreCase) {
            $imatches[$entry.Path] = $score
            if ($score -gt $iHiRank) {
                $iBestMatch = $entry.Path
                $iHiRank = $score
            }
        }
    }

    [pscustomobject]@{
        Pattern = $pattern
        Matches = $matches
        BestMatch = $bestMatch
        MatchCommon = Get-ZCommonPrefix -Matches $matches
        IMatches = $imatches
        IBestMatch = $iBestMatch
        IMatchCommon = Get-ZCommonPrefix -Matches $imatches
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
    Get-ZEntries | ForEach-Object {
        if ($_.Rank -ge 1) {
            if ($_.Path -eq $Path) {
                $entries[$_.Path] = @{
                    rank = $_.Rank + 1
                    time = $now
                }
            } else {
                $entries[$_.Path] = @{
                    rank = $_.Rank
                    time = $_.Time
                }
            }
            $totalCount += $_.Rank
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
    $literalArgs = $false
    
    # parse arguments
    for ($i = 0; $i -lt $rawArgs.Count; $i++) {
        $arg = $rawArgs[$i]
        if (-not $literalArgs -and $arg -eq '--') {
            $literalArgs = $true
        } elseif (
            -not $literalArgs -and
            $arg -match '^-' -and
            $arg.Length -gt 1 -and
            $arg.Substring(1) -match '^[cehlrtx]+$'
        ) {
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
        Remove-ZEntry -Path (Get-Location).Path
        return
    }
    
    if ($query.Count -eq 0) {
        $list = $true
    }
    
    # no datafile yet
    if (-not (Test-Path -LiteralPath $global:_Z_DATA -PathType Leaf)) { 
        Write-Host "No z data file found. cd around to build it up." -ForegroundColor Yellow
        return 
    }

    $found = $null
    try {
        $found = Find-ZMatches -Query $query -Type $type -RestrictPwd:$restrictPwd
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        return
    }
    
    # output results
    if ($list) {
        $matchTable = if ($found.Matches.Count -gt 0) {
            $found.Matches
        } else { 
            $found.IMatches
        }

        $sorted = $matchTable.GetEnumerator() | Sort-Object Value
        $sorted | ForEach-Object {
            Write-Host ("{0,-10} {1}" -f $_.Value, $_.Key)
        }
    } else {
        # navigate to best match
        $target = if ($found.BestMatch) {
            if ($type -eq 'frecent' -and $found.MatchCommon) { $found.MatchCommon } else { $found.BestMatch }
        } elseif ($found.IBestMatch) {
            if ($type -eq 'frecent' -and $found.IMatchCommon) { $found.IMatchCommon } else { $found.IBestMatch }
        } else {
            $null
        }
        
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
    
    if (-not (Test-Path -LiteralPath $global:_Z_DATA -PathType Leaf)) { return }
    
    $pattern = $wordToComplete -replace ' ', '.*'
    $dirs = @()
    
    Get-ZEntries | ForEach-Object {
        if ($_.Path -like "*$pattern*") {
            $dirs += $_.Path
        }
    }
    
    $dirs | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_)
    }
}
