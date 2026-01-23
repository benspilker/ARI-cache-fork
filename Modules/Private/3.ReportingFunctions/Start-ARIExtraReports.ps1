<#
.Synopsis
Module for Extra Reports

.DESCRIPTION
This script processes and creates additional report sheets such as Quotas, Security Center, Policies, and Advisory.

.Link
https://github.com/microsoft/ARI/Modules/Private/3.ReportingFunctions/Start-ARIExtraReports.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Start-ARIExtraReports {
    Param($File, $Quotas, $SecurityCenter, $SkipPolicy, $SkipAdvisory, $IncludeCosts, $TableStyle)

    Write-Progress -activity 'Azure Inventory' -Status "70% Complete." -PercentComplete 70 -CurrentOperation "Reporting Extra Resources.."

    <################################################ QUOTAS #######################################################>

    if(![string]::IsNullOrEmpty($Quotas))
        {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Generating Quota Usage Sheet.')
            Write-Progress -Id 1 -activity 'Azure Resource Inventory Quota Usage' -Status "50% Complete." -PercentComplete 50 -CurrentOperation "Building Quota Sheet"

            Build-ARIQuotaReport -File $File -AzQuota $Quotas -TableStyle $TableStyle

            Write-Progress -Id 1 -activity 'Azure Resource Inventory Quota Usage' -Status "100% Complete." -Completed
        }

    <################################################ SECURITY CENTER #######################################################>

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Checking if Should Generate Security Center Sheet.')
    if ($SecurityCenter.IsPresent) {
        if(get-job | Where-Object {$_.Name -eq 'Security'})
            {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Generating Security Center Sheet.')

                while (get-job -Name 'Security' | Where-Object { $_.State -eq 'Running' }) {
                    Write-Progress -Id 1 -activity 'Processing Security Center Advisories' -Status "50% Complete." -PercentComplete 50
                    Start-Sleep -Seconds 2
                }

                $Sec = Receive-Job -Name 'Security' -ErrorAction SilentlyContinue
                Remove-Job -Name 'Security' -ErrorAction SilentlyContinue | Out-Null
                
                # Ensure Sec is an array for safe handling
                if ($null -eq $Sec) {
                    $Sec = @()
                } elseif ($Sec -isnot [System.Array]) {
                    $Sec = @($Sec)
                }

                Build-ARISecCenterReport -File $File -Sec $Sec -TableStyle $TableStyle

                Write-Progress -Id 1 -activity 'Processing Security Center Advisories'  -Status "100% Complete." -Completed
            }

    }

    <################################################ POLICY #######################################################>

    # Receive Advisory job results BEFORE Policy cleanup removes all jobs
    # Handle both switch parameter and boolean value for SkipAdvisory
    $skipAdvisoryCheck = if ($SkipAdvisory -is [switch]) { $SkipAdvisory.IsPresent } else { $SkipAdvisory -eq $true }
    
    $Adv = $null
    if (-not $skipAdvisoryCheck) {
        $AdvisoryJob = Get-Job -Name 'Advisory' -ErrorAction SilentlyContinue
        if ($null -ne $AdvisoryJob) {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Receiving Advisory job results before Policy cleanup.')
            while (get-job -Name 'Advisory' | Where-Object { $_.State -eq 'Running' }) {
                Start-Sleep -Seconds 1
            }
            $Adv = Receive-Job -Name 'Advisory' -ErrorAction SilentlyContinue
            Remove-Job -Name 'Advisory' -ErrorAction SilentlyContinue | Out-Null
            # Ensure Adv is an array for safe handling
            if ($null -eq $Adv) {
                $Adv = @()
            } elseif ($Adv -isnot [System.Array]) {
                $Adv = @($Adv)
            }
        }
    }

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Checking if Should Generate Policy Sheet.')
    if (!$SkipPolicy.IsPresent) {
        if(get-job | Where-Object {$_.Name -eq 'Policy'})
            {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Generating Policy Sheet.')

                # Aggressive memory cleanup BEFORE receiving Policy job results to free memory from previous operations
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Running aggressive memory cleanup before receiving Policy job results.')
                try {
                    # Remove any other completed jobs first (Advisory was already received above)
                    Get-Job | Where-Object {$_.State -ne 'Running'} | Remove-Job -Force -ErrorAction SilentlyContinue
                    # Multiple aggressive GC collections
                    for ($i = 1; $i -le 5; $i++) {
                        [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $false)
                        [System.GC]::WaitForPendingFinalizers()
                        [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                    }
                    Clear-ARIMemory
                } catch {
                    Write-Debug "  Warning: Pre-Policy memory cleanup had issues: $_"
                }

                while (get-job -Name 'Policy' | Where-Object { $_.State -eq 'Running' }) {
                    Write-Progress -Id 1 -activity 'Processing Policies' -Status "50% Complete." -PercentComplete 50
                    Start-Sleep -Seconds 2
                }

                $Pol = Receive-Job -Name 'Policy' -ErrorAction SilentlyContinue
                Remove-Job -Name 'Policy' -ErrorAction SilentlyContinue | Out-Null
                
                # Ensure Pol is an array for safe handling
                if ($null -eq $Pol) {
                    $Pol = @()
                } elseif ($Pol -isnot [System.Array]) {
                    $Pol = @($Pol)
                }

                # Aggressive memory cleanup after receiving Policy data but before Excel generation
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Running aggressive memory cleanup before Policy sheet Excel generation.')
                try {
                    # Remove any remaining jobs
                    Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
                    # Multiple aggressive GC collections
                    for ($i = 1; $i -le 10; $i++) {
                        [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $false)
                        [System.GC]::WaitForPendingFinalizers()
                        [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                    }
                    Clear-ARIMemory
                } catch {
                    Write-Debug "  Warning: Memory cleanup had issues: $_"
                }

                Build-ARIPolicyReport -File $File -Pol $Pol -TableStyle $TableStyle
                
                # Cleanup after Policy sheet generation
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Running memory cleanup after Policy sheet generation.')
                try {
                    [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                    Clear-ARIMemory
                } catch {
                    Write-Debug "  Warning: Post-Policy cleanup had issues: $_"
                }

                Write-Progress -Id 1 -activity 'Processing Policies'  -Status "100% Complete." -Completed

                Start-Sleep -Milliseconds 200
            }
    }

    <################################################ ADVISOR #######################################################>

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Checking if Should Generate Advisory Sheet.')
    # Use the same skipAdvisoryCheck variable defined above
    if (-not $skipAdvisoryCheck) {
        # Advisory job results were already received before Policy cleanup (see above)
        if ($null -ne $Adv) {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Generating Advisor Sheet.')

            # Only generate sheet if we have Advisory data
            if ($Adv.Count -gt 0) {
                Build-ARIAdvisoryReport -File $File -Adv $Adv -TableStyle $TableStyle
                Write-Progress -Id 1 -activity 'Processing Advisories'  -Status "100% Complete." -Completed
            } else {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'No Advisory data to report - skipping Advisory sheet.')
            }

            Start-Sleep -Milliseconds 200
        } else {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'No Advisory data available - skipping Advisory sheet.')
        }
    }

    <################################################################### SUBSCRIPTIONS ###################################################################>

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Generating Subscription sheet.')

    Write-Progress -activity 'Azure Resource Inventory Subscriptions' -Status "50% Complete." -PercentComplete 50 -CurrentOperation "Building Subscriptions Sheet"

    $SubscriptionsJob = Get-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
    if ($null -ne $SubscriptionsJob) {
        while ($SubscriptionsJob | Where-Object { $_.State -eq 'Running' }) {
            Write-Progress -Id 1 -activity 'Processing Subscriptions' -Status "50% Complete." -PercentComplete 50
            Start-Sleep -Seconds 2
            $SubscriptionsJob = Get-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
        }

        # Check if job failed
        if ($SubscriptionsJob.State -eq 'Failed') {
            $jobError = Receive-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
            Write-Error "Subscriptions job failed: $($SubscriptionsJob | Format-List | Out-String)"
            if ($jobError) {
                Write-Error "Job error output: $($jobError | Out-String)"
            }
            $AzSubs = @()
        } else {
            try {
                $AzSubs = Receive-Job -Name 'Subscriptions' -ErrorAction Stop
            } catch {
                Write-Error "Error receiving Subscriptions job results: $($_.Exception.Message)"
                Write-Error "Stack trace: $($_.ScriptStackTrace)"
                $AzSubs = @()
            }
        }
        Remove-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue | Out-Null
        
        # Ensure AzSubs is an array for safe handling
        if ($null -eq $AzSubs) {
            $AzSubs = @()
        } elseif ($AzSubs -isnot [System.Array]) {
            # If it's a single object, wrap it in an array
            $AzSubs = @($AzSubs)
        }
    } else {
        Write-Debug "  Warning: Subscriptions job not found - initializing empty array"
        $AzSubs = @()
    }

    Build-ARISubsReport -File $File -Sub $AzSubs -IncludeCosts $IncludeCosts -TableStyle $TableStyle

    Clear-ARIMemory

    Write-Progress -activity 'Azure Resource Inventory Subscriptions' -Status "100% Complete." -Completed

    Write-Progress -activity 'Azure Inventory' -Status "80% Complete." -PercentComplete 80 -CurrentOperation "Completed Extra Resources Reporting.."
}