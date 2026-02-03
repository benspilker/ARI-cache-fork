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
    if ($null -eq $JobNames -or $JobNames.Count -eq 0) {
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
}
