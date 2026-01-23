<#
.Synopsis
Main module for Excel Report Building

.DESCRIPTION
This module is the main module for building the Excel Report.

.Link
https://github.com/microsoft/ARI/Modules/Private/0.MainFunctions/Start-ARIReporOrchestration.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola

#>
Function Start-ARIReporOrchestration {
    Param($ReportCache,
    $SecurityCenter,
    $File,
    $Quotas,
    $SkipPolicy,
    $SkipAdvisory,
    $Automation,
    $TableStyle,
    $IncludeCosts,
    $Advisories)

    Write-Progress -activity 'Azure Inventory' -Status "65% Complete." -PercentComplete 65 -CurrentOperation "Starting the Report Phase.."

    <############################################################## REPORT CREATION ###################################################################>

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Resource Reporting Cache.')
    Write-Host "[DEBUG] Start-ARIReporOrchestration: About to call Start-ARIExcelJob" -ForegroundColor Magenta
    try {
        Start-ARIExcelJob -ReportCache $ReportCache -TableStyle $TableStyle -File $File
        Write-Host "[DEBUG] Start-ARIReporOrchestration: Start-ARIExcelJob completed successfully" -ForegroundColor Magenta
    } catch {
        $errorMsg = "Error in Start-ARIExcelJob: $($_.Exception.Message)"
        $errorLine = if ($null -ne $_.InvocationInfo) { $_.InvocationInfo.ScriptLineNumber } else { "Unknown" }
        $errorFunc = if ($null -ne $_.InvocationInfo -and $null -ne $_.InvocationInfo.FunctionName) { $_.InvocationInfo.FunctionName } else { "Unknown" }
        $errorStack = if ($null -ne $_.ScriptStackTrace) { $_.ScriptStackTrace } else { "No stack trace available" }
        Write-Host "[ERROR] $errorMsg" -ForegroundColor Red
        Write-Host "[ERROR] Line: $errorLine, Function: $errorFunc" -ForegroundColor Red
        Write-Host "[ERROR] Stack: $errorStack" -ForegroundColor Red
        Write-Error $errorMsg
        Write-Error "Stack trace: $errorStack"
        throw
    }

    # Receive Subscriptions job results BEFORE memory cleanup removes all jobs
    $script:AzSubs = $null
    $SubscriptionsJob = Get-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
    if ($null -ne $SubscriptionsJob) {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Receiving Subscriptions job results before memory cleanup.')
        while ($SubscriptionsJob | Where-Object { $_.State -eq 'Running' }) {
            Start-Sleep -Seconds 1
            $SubscriptionsJob = Get-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
            if ($null -eq $SubscriptionsJob) { break }
        }
        if ($null -ne $SubscriptionsJob) {
            if ($SubscriptionsJob.State -eq 'Failed') {
                $jobError = Receive-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job failed: ' + ($jobError | Out-String))
                $script:AzSubs = @()
            } else {
                $script:AzSubs = Receive-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
                if ($null -eq $script:AzSubs) {
                    $script:AzSubs = @()
                } elseif ($script:AzSubs -isnot [System.Array]) {
                    $script:AzSubs = @($script:AzSubs)
                }
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions data received and stored in script scope: Count=' + $script:AzSubs.Count)
            }
            Remove-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue | Out-Null
        }
    }

    # Aggressive memory cleanup between Excel phases to prevent OOM
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Running aggressive memory cleanup after Start-ARIExcelJob.')
    try {
        Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
        for ($i = 1; $i -le 5; $i++) {
            [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $false)
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
        }
        Clear-ARIMemory
    } catch {
        Write-Debug "  Warning: Memory cleanup after Start-ARIExcelJob had issues: $_"
    }

    <############################################################## REPORT EXTRA DETAILS ###################################################################>

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Reporting Extra Details.')
    Write-Host "[DEBUG] Start-ARIReporOrchestration: About to call Start-ARIExcelExtraData" -ForegroundColor Magenta
    try {
        Start-ARIExcelExtraData -File $File
        Write-Host "[DEBUG] Start-ARIReporOrchestration: Start-ARIExcelExtraData completed successfully" -ForegroundColor Magenta
    } catch {
        $errorMsg = "Error in Start-ARIExcelExtraData: $($_.Exception.Message)"
        $errorLine = if ($null -ne $_.InvocationInfo) { $_.InvocationInfo.ScriptLineNumber } else { "Unknown" }
        $errorFunc = if ($null -ne $_.InvocationInfo -and $null -ne $_.InvocationInfo.FunctionName) { $_.InvocationInfo.FunctionName } else { "Unknown" }
        $errorStack = if ($null -ne $_.ScriptStackTrace) { $_.ScriptStackTrace } else { "No stack trace available" }
        Write-Host "[ERROR] $errorMsg" -ForegroundColor Red
        Write-Host "[ERROR] Line: $errorLine, Function: $errorFunc" -ForegroundColor Red
        Write-Host "[ERROR] Stack: $errorStack" -ForegroundColor Red
        Write-Error $errorMsg
        Write-Error "Stack trace: $errorStack"
        throw
    }

    # Aggressive memory cleanup between Excel phases to prevent OOM
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Running aggressive memory cleanup after Start-ARIExcelExtraData.')
    try {
        Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
        for ($i = 1; $i -le 5; $i++) {
            [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $false)
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
        }
        Clear-ARIMemory
    } catch {
        Write-Debug "  Warning: Memory cleanup after Start-ARIExcelExtraData had issues: $_"
    }

    <############################################################## EXTRA REPORTS ###################################################################>

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Default Data Reporting.')
    Write-Host "[DEBUG] Start-ARIReporOrchestration: About to call Start-ARIExtraReports" -ForegroundColor Magenta
    try {
        Start-ARIExtraReports -File $File -Quotas $Quotas -SecurityCenter $SecurityCenter -SkipPolicy $SkipPolicy -SkipAdvisory $SkipAdvisory -IncludeCosts $IncludeCosts -TableStyle $TableStyle -Advisories $Advisories
        Write-Host "[DEBUG] Start-ARIReporOrchestration: Start-ARIExtraReports completed successfully" -ForegroundColor Magenta
    } catch {
        $errorMsg = "Error in Start-ARIExtraReports: $($_.Exception.Message)"
        $errorLine = if ($null -ne $_.InvocationInfo) { $_.InvocationInfo.ScriptLineNumber } else { "Unknown" }
        $errorFunc = if ($null -ne $_.InvocationInfo -and $null -ne $_.InvocationInfo.FunctionName) { $_.InvocationInfo.FunctionName } else { "Unknown" }
        $errorStack = if ($null -ne $_.ScriptStackTrace) { $_.ScriptStackTrace } else { "No stack trace available" }
        Write-Host "[ERROR] $errorMsg" -ForegroundColor Red
        Write-Host "[ERROR] Line: $errorLine, Function: $errorFunc" -ForegroundColor Red
        Write-Host "[ERROR] Stack: $errorStack" -ForegroundColor Red
        Write-Error $errorMsg
        Write-Error "Stack trace: $errorStack"
        throw
    }

}