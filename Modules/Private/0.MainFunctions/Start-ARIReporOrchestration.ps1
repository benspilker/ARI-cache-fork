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
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job state: ' + $SubscriptionsJob.State)
        
        # Wait for job to complete with timeout
        $maxWaitTime = 60  # Maximum wait time in seconds
        $waitTime = 0
        while ($SubscriptionsJob.State -eq 'Running' -and $waitTime -lt $maxWaitTime) {
            Start-Sleep -Seconds 2
            $waitTime += 2
            $SubscriptionsJob = Get-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
            if ($null -eq $SubscriptionsJob) { 
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job disappeared while waiting')
                break 
            }
            if ($waitTime % 10 -eq 0) {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Still waiting for Subscriptions job to complete... (waited ' + $waitTime + 's)')
            }
        }
        
        if ($null -ne $SubscriptionsJob) {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job final state: ' + $SubscriptionsJob.State)
            
            if ($SubscriptionsJob.State -eq 'Failed') {
                $jobError = Receive-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job failed: ' + ($jobError | Out-String))
                if ($SubscriptionsJob | Get-Member -Name 'Error' -ErrorAction SilentlyContinue) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job error details: ' + ($SubscriptionsJob.Error | Out-String))
                }
                $script:AzSubs = @()
            } elseif ($SubscriptionsJob.State -eq 'Completed') {
                # Get all output from the job (including debug messages)
                $jobOutput = Receive-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
                
                if ($null -ne $jobOutput) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job output type: ' + $jobOutput.GetType().FullName)
                    if ($jobOutput -is [System.Array]) {
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job output is array with ' + $jobOutput.Count + ' element(s)')
                        
                        # Less aggressive filtering: Only filter out DebugRecord objects, keep everything else
                        # PSCustomObjects (the actual data) will be preserved
                        # Strings that are NOT debug messages will also be preserved (though unlikely)
                        $dataObjects = $jobOutput | Where-Object { $_ -isnot [System.Management.Automation.DebugRecord] }
                        
                        # Further filter: Keep only PSCustomObjects (the actual subscription data)
                        # This ensures we only keep data objects, not any stray strings
                        $psCustomObjects = $dataObjects | Where-Object { $_ -is [PSCustomObject] }
                        
                        if ($psCustomObjects.Count -gt 0) {
                            $script:AzSubs = $psCustomObjects
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Filtered to ' + $psCustomObjects.Count + ' PSCustomObject(s) from job output')
                        } elseif ($dataObjects.Count -gt 0) {
                            # If we have data objects but they're not PSCustomObjects, check their type
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Found ' + $dataObjects.Count + ' non-DebugRecord object(s), but none are PSCustomObjects')
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'First object type: ' + $dataObjects[0].GetType().FullName)
                            # Try to use them anyway - might be deserialized objects
                            $script:AzSubs = $dataObjects
                        } else {
                            # Check what we filtered out
                            $debugRecordCount = ($jobOutput | Where-Object { $_ -is [System.Management.Automation.DebugRecord] }).Count
                            $stringCount = ($jobOutput | Where-Object { $_ -is [string] }).Count
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Job output contains ' + $debugRecordCount + ' DebugRecord(s), ' + $stringCount + ' string(s), 0 data objects')
                            $script:AzSubs = @()
                        }
                    } else {
                        # Single object - check if it's a PSCustomObject (data) or something else
                        if ($jobOutput -is [PSCustomObject]) {
                            $script:AzSubs = $jobOutput
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job returned single PSCustomObject')
                        } elseif ($jobOutput -isnot [System.Management.Automation.DebugRecord] -and $jobOutput -isnot [string]) {
                            # Might be a deserialized object - try to use it
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job returned single object of type: ' + $jobOutput.GetType().FullName)
                            $script:AzSubs = $jobOutput
                        } else {
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job returned string/debug record instead of data')
                            $script:AzSubs = @()
                        }
                    }
                } else {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job output is null')
                    $script:AzSubs = @()
                }
                
                # Ensure AzSubs is an array
                if ($null -eq $script:AzSubs) {
                    $script:AzSubs = @()
                } elseif ($script:AzSubs -isnot [System.Array]) {
                    $script:AzSubs = @($script:AzSubs)
                }
                
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions data received and stored in script scope: Count=' + $script:AzSubs.Count)
                if ($script:AzSubs.Count -eq 0 -and $null -ne $jobOutput) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Warning: Subscriptions job completed but returned 0 results.')
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Job output sample (first 500 chars): ' + ($jobOutput | Out-String -Width 200 | Select-Object -First 500))
                }
            } else {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job in unexpected state: ' + $SubscriptionsJob.State)
                $script:AzSubs = @()
            }
            Remove-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue | Out-Null
        }
    } else {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job not found when trying to receive results.')
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