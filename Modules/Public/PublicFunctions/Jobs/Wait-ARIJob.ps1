<#
.Synopsis
Wait for ARI Jobs to Complete

.DESCRIPTION
This script waits for the completion of specified ARI jobs.

.Link
https://github.com/microsoft/ARI/Modules/Public/PublicFunctions/Jobs/Wait-ARIJob.ps1

.COMPONENT
    This powershell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola

#>
function Wait-ARIJob {
    Param($JobNames, $JobType, $LoopTime)

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Jobs Collector.')

    # Normalize and guard against null/empty job name lists
    if ($null -eq $JobNames) {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'No job names provided to Jobs Collector; skipping wait.')
        return
    }
    if ($JobNames -isnot [System.Array]) {
        $JobNames = @($JobNames)
    }
    $JobNames = $JobNames | Where-Object { $_ -and -not [string]::IsNullOrEmpty($_) }
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Jobs Collector received job names: '+($JobNames -join ', '))
    $jobNamesCount = ($JobNames | Measure-Object).Count
    if ($null -eq $JobNames -or $jobNamesCount -eq 0) {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'No valid job names provided to Jobs Collector; skipping wait.')
        return
    }

    $c = 0

    try {
        while (get-job -Name $JobNames | Where-Object { $_.State -eq 'Running' }) {
            $jb = get-job -Name $JobNames
        # Ensure $jb is always an array for safe .Count access
        if ($jb -isnot [System.Array]) {
            $jb = @($jb)
        }
        
        # Safely get running jobs count
        $runningJobs = $jb | Where-Object { $_.State -eq 'Running' }
        if ($runningJobs -isnot [System.Array]) {
            $runningJobs = @($runningJobs)
        }

        # Use Measure-Object to avoid .Count on non-collection types
        $jbCount = ($jb | Measure-Object).Count
        $runningCount = ($runningJobs | Measure-Object).Count

        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Jobs Collector types: JobNames=' + $JobNames.GetType().FullName + '; Jobs=' + ($jb | Select-Object -First 1 | ForEach-Object { $_.GetType().FullName }) + '; RunningJobs=' + ($runningJobs | Select-Object -First 1 | ForEach-Object { $_.GetType().FullName }))
        
        if ($jbCount -gt 0) {
            $c = ((($jbCount - $runningCount) / $jbCount) * 100)
        } else {
            $c = 100
        }
        
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+"$JobType Jobs Still Running: "+[string]$runningCount)
            $c = [math]::Round($c)
            Write-Progress -Id 1 -activity "Processing $JobType Jobs" -Status "$c% Complete." -PercentComplete $c
            Start-Sleep -Seconds $LoopTime
        }
    } catch {
        Write-Error "Wait-ARIJob error: $($_.Exception.Message)"
        Write-Error "Wait-ARIJob line: $($_.InvocationInfo.ScriptLineNumber)"
        Write-Error "Wait-ARIJob stack: $($_.ScriptStackTrace)"
        throw
    }
    Write-Progress -Id 1 -activity "Processing $JobType Jobs" -Status "100% Complete." -Completed

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Jobs Complete.')

    # [PATCHED-PER-MODULE-GUARD] Cgroup memory check between modules (85% threshold)
    # Legacy marker retained for the existing CSP idempotency check. The active cutoff is 75%.
    # Resolve the process membership and mount rather than assuming cgroup files live
    # at /sys/fs/cgroup; Windmill's unshare namespace may expose a nested mount/path.
    $perModuleGuardMaxPct = 75
    $perModuleUsedBytes = $null
    $perModuleLimitBytes = $null
    $perModulePathUsed = 'unavailable'
    $perModuleBestPct = $null
    $perModuleCandidateDirs = @()
    $perModuleGroupV2 = $null
    $perModuleGroupV1Memory = $null

    try {
        foreach ($line in [System.IO.File]::ReadAllLines('/proc/self/cgroup')) {
            $fields = $line -split ':', 3
            if ($fields.Count -lt 3) { continue }
            if ($fields[0] -eq '0' -and [string]::IsNullOrEmpty($fields[1])) {
                $perModuleGroupV2 = $fields[2]
            } elseif (@($fields[1] -split ',') -contains 'memory') {
                $perModuleGroupV1Memory = $fields[2]
            }
        }
    } catch { }

    try {
        foreach ($line in [System.IO.File]::ReadAllLines('/proc/self/mountinfo')) {
            $sections = $line -split ' - ', 2
            if ($sections.Count -lt 2) { continue }
            $mountFields = @($sections[0].Trim() -split '\s+')
            $fsFields = @($sections[1].Trim() -split '\s+')
            if ($mountFields.Count -lt 5 -or $fsFields.Count -lt 3) { continue }

            $fsType = $fsFields[0]
            $mountRoot = $mountFields[3].Replace('\040', ' ').Replace('\011', "`t").Replace('\012', "`n").Replace('\134', '\')
            $mountPoint = $mountFields[4].Replace('\040', ' ').Replace('\011', "`t").Replace('\012', "`n").Replace('\134', '\')
            $memberPath = $null
            if ($fsType -eq 'cgroup2' -and $null -ne $perModuleGroupV2) {
                $memberPath = $perModuleGroupV2
            } elseif ($fsType -eq 'cgroup' -and $null -ne $perModuleGroupV1Memory -and (($fsFields -contains 'memory') -or ($fsFields[-1] -match '(^|,)memory(,|$)'))) {
                $memberPath = $perModuleGroupV1Memory
            }
            if ($null -eq $memberPath) { continue }

            $rootPrefix = $mountRoot.TrimEnd('/')
            if ([string]::IsNullOrEmpty($rootPrefix)) {
                $relativePath = $memberPath.TrimStart('/')
            } elseif ($memberPath -eq $mountRoot) {
                $relativePath = ''
            } elseif ($memberPath.StartsWith($rootPrefix + '/')) {
                $relativePath = $memberPath.Substring($rootPrefix.Length).TrimStart('/')
            } else {
                $relativePath = $null
            }
            if ($null -ne $relativePath) {
                if ([string]::IsNullOrEmpty($relativePath)) {
                    $perModuleCandidateDirs += $mountPoint
                } else {
                    $perModuleCandidateDirs += (Join-Path $mountPoint $relativePath)
                }
            }
            $perModuleCandidateDirs += $mountPoint
        }
    } catch { }

    # If mountinfo is hidden but cgroup membership is readable, combine the
    # membership path with common v2/v1 mount roots as a second discovery path.
    if ($null -ne $perModuleGroupV2) {
        $v2Relative = $perModuleGroupV2.TrimStart('/')
        foreach ($v2Base in @('/sys/fs/cgroup', '/sys/fs/cgroup/unified')) {
            if ([string]::IsNullOrEmpty($v2Relative)) {
                $perModuleCandidateDirs += $v2Base
            } else {
                $perModuleCandidateDirs += (Join-Path $v2Base $v2Relative)
            }
        }
    }
    if ($null -ne $perModuleGroupV1Memory) {
        $v1Relative = $perModuleGroupV1Memory.TrimStart('/')
        if ([string]::IsNullOrEmpty($v1Relative)) {
            $perModuleCandidateDirs += '/sys/fs/cgroup/memory'
        } else {
            $perModuleCandidateDirs += (Join-Path '/sys/fs/cgroup/memory' $v1Relative)
        }
    }
    # Final fallbacks for common cgroup mount roots.
    $perModuleCandidateDirs += @('/sys/fs/cgroup', '/sys/fs/cgroup/unified', '/sys/fs/cgroup/memory')
    $perModuleProbeDirs = @()
    foreach ($candidate in @($perModuleCandidateDirs | Select-Object -Unique)) {
        $currentDir = $candidate
        for ($depth = 0; $depth -lt 16 -and -not [string]::IsNullOrWhiteSpace($currentDir); $depth++) {
            if (-not ($perModuleProbeDirs -contains $currentDir)) { $perModuleProbeDirs += $currentDir }
            try { $parentDir = [System.IO.Directory]::GetParent($currentDir) } catch { $parentDir = $null }
            if ($null -eq $parentDir -or $parentDir.FullName -eq $currentDir) { break }
            $currentDir = $parentDir.FullName
        }
    }

    # Check every matching ancestor and enforce the tightest observed finite limit.
    # memory.current at a parent cgroup includes usage from its child cgroups.
    foreach ($cgroupDir in $perModuleProbeDirs) {
        $v2UsedPath = Join-Path $cgroupDir 'memory.current'
        $v2LimitPath = Join-Path $cgroupDir 'memory.max'
        if ((Test-Path -LiteralPath $v2UsedPath) -and (Test-Path -LiteralPath $v2LimitPath)) {
            try {
                $usedText = [System.IO.File]::ReadAllText($v2UsedPath).Trim()
                $limitText = [System.IO.File]::ReadAllText($v2LimitPath).Trim()
                if ($usedText -match '^\d+$' -and $limitText -match '^\d+$') {
                    $candidateUsed = [int64]$usedText
                    $candidateLimit = [int64]$limitText
                    if ($candidateLimit -gt 0 -and $candidateLimit -lt 1099511627776) {
                        $candidatePct = [math]::Round((([double]$candidateUsed / [double]$candidateLimit) * 100), 1)
                        if ($null -eq $perModuleBestPct -or $candidatePct -gt $perModuleBestPct) {
                            $perModuleUsedBytes = $candidateUsed
                            $perModuleLimitBytes = $candidateLimit
                            $perModuleBestPct = $candidatePct
                            $perModulePathUsed = "v2:$cgroupDir"
                        }
                    }
                } elseif ($limitText -eq 'max' -and $perModulePathUsed -eq 'unavailable') {
                    $perModulePathUsed = "v2-unlimited:$cgroupDir"
                }
            } catch { }
        }

        $v1UsedPath = Join-Path $cgroupDir 'memory.usage_in_bytes'
        $v1LimitPath = Join-Path $cgroupDir 'memory.limit_in_bytes'
        if ((Test-Path -LiteralPath $v1UsedPath) -and (Test-Path -LiteralPath $v1LimitPath)) {
            try {
                $usedText = [System.IO.File]::ReadAllText($v1UsedPath).Trim()
                $limitText = [System.IO.File]::ReadAllText($v1LimitPath).Trim()
                if ($usedText -match '^\d+$' -and $limitText -match '^\d+$') {
                    $candidateUsed = [int64]$usedText
                    $candidateLimit = [int64]$limitText
                    # cgroup v1 commonly represents "unlimited" with a near-Int64-max value.
                    if ($candidateLimit -gt 0 -and $candidateLimit -lt 1099511627776) {
                        $candidatePct = [math]::Round((([double]$candidateUsed / [double]$candidateLimit) * 100), 1)
                        if ($null -eq $perModuleBestPct -or $candidatePct -gt $perModuleBestPct) {
                            $perModuleUsedBytes = $candidateUsed
                            $perModuleLimitBytes = $candidateLimit
                            $perModuleBestPct = $candidatePct
                            $perModulePathUsed = "v1:$cgroupDir"
                        }
                    } elseif ($perModulePathUsed -eq 'unavailable') {
                        $perModulePathUsed = "v1-unlimited:$cgroupDir"
                    }
                }
            } catch { }
        }
    }

    if ($null -ne $perModuleBestPct -and $null -ne $perModuleUsedBytes -and $null -ne $perModuleLimitBytes) {
        $perModuleUsedMB = [math]::Round($perModuleUsedBytes / 1MB, 1)
        $perModuleLimitMB = [math]::Round($perModuleLimitBytes / 1MB, 1)
        Write-Host "  [PER-MODULE-GUARD] cgroup path=$perModulePathUsed used=$perModuleUsedMB MB limit=$perModuleLimitMB MB pct=$perModuleBestPct% threshold=$perModuleGuardMaxPct%" -ForegroundColor DarkGray
        if ($perModuleBestPct -gt $perModuleGuardMaxPct) {
            $guardMsg = '[PATCHED-PER-MODULE-GUARD] Aborting module loop: cgroup at ' + $perModuleBestPct + '% (' + $perModuleUsedMB + 'MB / ' + $perModuleLimitMB + 'MB) > ' + $perModuleGuardMaxPct + '% threshold.'
            Write-Host "  $guardMsg" -ForegroundColor Yellow
            try { Set-Content -Path '/tmp/per_module_oom_guard.flag' -Value $guardMsg -ErrorAction SilentlyContinue } catch { }
            exit 42
        }
    } else {
        Write-Host "  [PER-MODULE-GUARD] cgroup memory limit unavailable (pathUsed=$perModulePathUsed, usedBytes=$perModuleUsedBytes, limitBytes=$perModuleLimitBytes) - threshold NOT enforced this check." -ForegroundColor Yellow
    }
}
