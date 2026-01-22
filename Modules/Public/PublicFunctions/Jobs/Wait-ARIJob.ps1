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

    $c = 0

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
        
        $jbCount = if ($null -ne $jb) { $jb.Count } else { 0 }
        $runningCount = if ($null -ne $runningJobs) { $runningJobs.Count } else { 0 }
        
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
    Write-Progress -Id 1 -activity "Processing $JobType Jobs" -Status "100% Complete." -Completed

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Jobs Complete.')
}