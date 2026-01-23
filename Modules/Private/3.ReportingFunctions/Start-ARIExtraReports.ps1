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
    Param($File, $Quotas, $SecurityCenter, $SkipPolicy, $SkipAdvisory, $IncludeCosts, $TableStyle, $Advisories)

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
    
    # Store Advisory data in script scope to prevent it from being cleared by memory cleanup
    $script:Adv = $null
    if (-not $skipAdvisoryCheck) {
        $AdvisoryJob = Get-Job -Name 'Advisory' -ErrorAction SilentlyContinue
        if ($null -ne $AdvisoryJob) {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Receiving Advisory job results before Policy cleanup.')
            while (get-job -Name 'Advisory' | Where-Object { $_.State -eq 'Running' }) {
                Start-Sleep -Seconds 1
            }
            $script:Adv = Receive-Job -Name 'Advisory' -ErrorAction SilentlyContinue
            Remove-Job -Name 'Advisory' -ErrorAction SilentlyContinue | Out-Null
            # Ensure Adv is an array for safe handling
            if ($null -eq $script:Adv) {
                $script:Adv = @()
            } elseif ($script:Adv -isnot [System.Array]) {
                $script:Adv = @($script:Adv)
            }
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisory data received and stored in script scope: Count=' + $script:Adv.Count)
        } else {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisory job not found when trying to receive results.')
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
        # Advisory job results were already received before Policy cleanup and stored in script scope (see above)
        # Use script-scoped variable to ensure it wasn't cleared by memory cleanup
        $Adv = $script:Adv
        if ($null -eq $Adv) {
            # Fallback: try to receive from job if script-scoped variable is null
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Script-scoped Advisory data is null - checking for Advisory job...')
            $AdvisoryJobCheck = Get-Job -Name 'Advisory' -ErrorAction SilentlyContinue
            if ($null -ne $AdvisoryJobCheck) {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisory job found: State=' + $AdvisoryJobCheck.State)
                # Try to receive it now
                $Adv = Receive-Job -Name 'Advisory' -ErrorAction SilentlyContinue
                if ($null -ne $Adv) {
                    # Store in script scope for safety
                    $script:Adv = $Adv
                    if ($script:Adv -isnot [System.Array]) {
                        $script:Adv = @($script:Adv)
                    }
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Received Advisory data from job: Count=' + $script:Adv.Count)
                }
                Remove-Job -Name 'Advisory' -ErrorAction SilentlyContinue | Out-Null
            } else {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisory job not found - may have completed and been removed already.')
            }
        }
        
        # Fallback: if still no data, process Advisories parameter directly
        if ($null -eq $Adv -and $null -ne $Advisories) {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisory job not available - processing Advisories parameter directly...')
            # Ensure Advisories is an array
            $advisoriesArray = if ($null -eq $Advisories) {
                @()
            } elseif ($Advisories -isnot [System.Array]) {
                @($Advisories)
            } else {
                $Advisories
            }
            
            if ($advisoriesArray.Count -gt 0) {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Processing Advisories directly using Start-ARIAdvisoryJob (Count=' + $advisoriesArray.Count + ')')
                # Load Start-ARIAdvisoryJob function if available
                # Path: From Modules/Private/3.ReportingFunctions/ to Modules/Public/PublicFunctions/Jobs/
                # Use same method as Start-ARIExcelJob.ps1 for reliable path resolution
                try {
                    $ParentPath = (Get-Item $PSScriptRoot).Parent.Parent
                    $ariModulePath = Join-Path (Join-Path $ParentPath "Public\PublicFunctions\Jobs") "Start-ARIAdvisoryJob.ps1"
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Attempting to load Start-ARIAdvisoryJob from: ' + $ariModulePath)
                    
                    if (Test-Path $ariModulePath) {
                        try {
                            . $ariModulePath
                            $Adv = Start-ARIAdvisoryJob -Advisories $advisoriesArray
                            # Ensure Adv is an array
                            if ($null -eq $Adv) {
                                $Adv = @()
                            } elseif ($Adv -isnot [System.Array]) {
                                $Adv = @($Adv)
                            }
                            # Store in script scope
                            $script:Adv = $Adv
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Processed Advisory data directly: Count=' + $Adv.Count)
                        } catch {
                            $loadError = $_.Exception.Message
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error loading/executing Start-ARIAdvisoryJob: ' + $loadError)
                            # Try to process inline if function loading fails
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Attempting inline Advisory processing...')
                            $Adv = @()
                            foreach ($advItem in $advisoriesArray) {
                                try {
                                    if ($null -ne $advItem -and $null -ne $advItem.PROPERTIES) {
                                        $data = $advItem.PROPERTIES
                                        if ($null -ne $data -and $null -ne $data.resourceMetadata -and $null -ne $data.resourceMetadata.resourceId) {
                                            $Savings = if([string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){0}Else{$data.extendedProperties.annualSavingsAmount}
                                            $SavingsCurrency = if([string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){'USD'}Else{$data.extendedProperties.savingsCurrency}
                                            $Resource = $data.resourceMetadata.resourceId.split('/')
                                            # Add bounds checking for array access
                                            if ($Resource.Count -lt 4) {
                                                $ResourceType = $data.impactedField
                                                $ResourceName = $data.impactedValue
                                                $Subscription = ''
                                                $ResourceGroup = ''
                                            } elseif ($Resource.Count -lt 5) {
                                                $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                                                $ResourceGroup = ''
                                                $ResourceType = $data.impactedField
                                                $ResourceName = $data.impactedValue
                                            } elseif ($Resource.Count -lt 9) {
                                                $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                                                $ResourceGroup = if ($Resource.Count -gt 4) { $Resource[4] } else { '' }
                                                $ResourceType = $data.impactedField
                                                $ResourceName = $data.impactedValue
                                            } else {
                                                $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                                                $ResourceGroup = if ($Resource.Count -gt 4) { $Resource[4] } else { '' }
                                                $ResourceType = if ($Resource.Count -gt 7) { ($Resource[6] + '/' + $Resource[7]) } else { $data.impactedField }
                                                $ResourceName = if ($Resource.Count -gt 8) { $Resource[8] } else { $data.impactedValue }
                                            }
                                            if ($null -ne $data.impactedField -and $data.impactedField -eq $ResourceType) {
                                                $ImpactedField = ''
                                            } else {
                                                $ImpactedField = $data.impactedField
                                            }
                                            if ($null -ne $data.impactedValue -and $data.impactedValue -eq $ResourceName) {
                                                $ImpactedValue = ''
                                            } else {
                                                $ImpactedValue = $data.impactedValue
                                            }
                                            $obj = @{
                                                'Subscription'           = $Subscription;
                                                'Resource Group'         = $ResourceGroup;
                                                'Resource Type'          = $ResourceType;
                                                'Name'                   = $ResourceName;
                                                'Detailed Type'          = $ImpactedField;
                                                'Detailed Name'          = $ImpactedValue;
                                                'Category'               = $data.category;
                                                'Impact'                 = $data.impact;
                                                'Description'            = if ($null -ne $data.shortDescription) { $data.shortDescription.problem } else { '' };
                                                'SKU'                    = $data.extendedProperties.sku;
                                                'Term'                   = $data.extendedProperties.term;
                                                'Look-back Period'       = $data.extendedProperties.lookbackPeriod;
                                                'Quantity'               = $data.extendedProperties.qty;
                                                'Savings Currency'       = $SavingsCurrency;
                                                'Annual Savings'         = "=$Savings";
                                                'Savings Region'         = $data.extendedProperties.region
                                            }
                                            $Adv += $obj
                                        }
                                    }
                                } catch {
                                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error processing Advisory item: ' + $_.Exception.Message)
                                    continue
                                }
                            }
                            $script:Adv = $Adv
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Processed Advisory data inline: Count=' + $Adv.Count)
                        }
                    } else {
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Warning: Start-ARIAdvisoryJob.ps1 not found at: ' + $ariModulePath)
                        # Fallback to inline processing
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Attempting inline Advisory processing...')
                        $Adv = @()
                        foreach ($advItem in $advisoriesArray) {
                            try {
                                if ($null -ne $advItem -and $null -ne $advItem.PROPERTIES) {
                                    $data = $advItem.PROPERTIES
                                    if ($null -ne $data -and $null -ne $data.resourceMetadata -and $null -ne $data.resourceMetadata.resourceId) {
                                        $Savings = if([string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){0}Else{$data.extendedProperties.annualSavingsAmount}
                                        $SavingsCurrency = if([string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){'USD'}Else{$data.extendedProperties.savingsCurrency}
                                        $Resource = $data.resourceMetadata.resourceId.split('/')
                                        # Add bounds checking for array access
                                        if ($Resource.Count -lt 4) {
                                            $ResourceType = $data.impactedField
                                            $ResourceName = $data.impactedValue
                                            $Subscription = ''
                                            $ResourceGroup = ''
                                        } elseif ($Resource.Count -lt 5) {
                                            $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                                            $ResourceGroup = ''
                                            $ResourceType = $data.impactedField
                                            $ResourceName = $data.impactedValue
                                        } elseif ($Resource.Count -lt 9) {
                                            $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                                            $ResourceGroup = if ($Resource.Count -gt 4) { $Resource[4] } else { '' }
                                            $ResourceType = $data.impactedField
                                            $ResourceName = $data.impactedValue
                                        } else {
                                            $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                                            $ResourceGroup = if ($Resource.Count -gt 4) { $Resource[4] } else { '' }
                                            $ResourceType = if ($Resource.Count -gt 7) { ($Resource[6] + '/' + $Resource[7]) } else { $data.impactedField }
                                            $ResourceName = if ($Resource.Count -gt 8) { $Resource[8] } else { $data.impactedValue }
                                        }
                                        if ($null -ne $data.impactedField -and $data.impactedField -eq $ResourceType) {
                                            $ImpactedField = ''
                                        } else {
                                            $ImpactedField = $data.impactedField
                                        }
                                        if ($null -ne $data.impactedValue -and $data.impactedValue -eq $ResourceName) {
                                            $ImpactedValue = ''
                                        } else {
                                            $ImpactedValue = $data.impactedValue
                                        }
                                        $obj = @{
                                            'Subscription'           = $Subscription;
                                            'Resource Group'         = $ResourceGroup;
                                            'Resource Type'          = $ResourceType;
                                            'Name'                   = $ResourceName;
                                            'Detailed Type'          = $ImpactedField;
                                            'Detailed Name'          = $ImpactedValue;
                                            'Category'               = $data.category;
                                            'Impact'                 = $data.impact;
                                            'Description'            = if ($null -ne $data.shortDescription) { $data.shortDescription.problem } else { '' };
                                            'SKU'                    = $data.extendedProperties.sku;
                                            'Term'                   = $data.extendedProperties.term;
                                            'Look-back Period'       = $data.extendedProperties.lookbackPeriod;
                                            'Quantity'               = $data.extendedProperties.qty;
                                            'Savings Currency'       = $SavingsCurrency;
                                            'Annual Savings'         = "=$Savings";
                                            'Savings Region'         = $data.extendedProperties.region
                                        }
                                        $Adv += $obj
                                    }
                                }
                            } catch {
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error processing Advisory item: ' + $_.Exception.Message)
                                continue
                            }
                        }
                        $script:Adv = $Adv
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Processed Advisory data inline: Count=' + $Adv.Count)
                    }
                } catch {
                    $pathError = $_.Exception.Message
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error resolving path to Start-ARIAdvisoryJob: ' + $pathError)
                    # Last resort: inline processing
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Attempting inline Advisory processing as fallback...')
                    $Adv = @()
                    foreach ($advItem in $advisoriesArray) {
                        try {
                            if ($null -ne $advItem -and $null -ne $advItem.PROPERTIES) {
                                $data = $advItem.PROPERTIES
                                if ($null -ne $data -and $null -ne $data.resourceMetadata -and $null -ne $data.resourceMetadata.resourceId) {
                                    $Savings = if([string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){0}Else{$data.extendedProperties.annualSavingsAmount}
                                    $SavingsCurrency = if([string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){'USD'}Else{$data.extendedProperties.savingsCurrency}
                                    $Resource = $data.resourceMetadata.resourceId.split('/')
                                    # Add bounds checking for array access
                                    if ($Resource.Count -lt 4) {
                                        $ResourceType = $data.impactedField
                                        $ResourceName = $data.impactedValue
                                        $Subscription = ''
                                        $ResourceGroup = ''
                                    } elseif ($Resource.Count -lt 5) {
                                        $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                                        $ResourceGroup = ''
                                        $ResourceType = $data.impactedField
                                        $ResourceName = $data.impactedValue
                                    } elseif ($Resource.Count -lt 9) {
                                        $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                                        $ResourceGroup = if ($Resource.Count -gt 4) { $Resource[4] } else { '' }
                                        $ResourceType = $data.impactedField
                                        $ResourceName = $data.impactedValue
                                    } else {
                                        $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                                        $ResourceGroup = if ($Resource.Count -gt 4) { $Resource[4] } else { '' }
                                        $ResourceType = if ($Resource.Count -gt 7) { ($Resource[6] + '/' + $Resource[7]) } else { $data.impactedField }
                                        $ResourceName = if ($Resource.Count -gt 8) { $Resource[8] } else { $data.impactedValue }
                                    }
                                    if ($null -ne $data.impactedField -and $data.impactedField -eq $ResourceType) {
                                        $ImpactedField = ''
                                    } else {
                                        $ImpactedField = $data.impactedField
                                    }
                                    if ($null -ne $data.impactedValue -and $data.impactedValue -eq $ResourceName) {
                                        $ImpactedValue = ''
                                    } else {
                                        $ImpactedValue = $data.impactedValue
                                    }
                                    $obj = @{
                                        'Subscription'           = $Subscription;
                                        'Resource Group'         = $ResourceGroup;
                                        'Resource Type'          = $ResourceType;
                                        'Name'                   = $ResourceName;
                                        'Detailed Type'          = $ImpactedField;
                                        'Detailed Name'          = $ImpactedValue;
                                        'Category'               = $data.category;
                                        'Impact'                 = $data.impact;
                                        'Description'            = if ($null -ne $data.shortDescription) { $data.shortDescription.problem } else { '' };
                                        'SKU'                    = $data.extendedProperties.sku;
                                        'Term'                   = $data.extendedProperties.term;
                                        'Look-back Period'       = $data.extendedProperties.lookbackPeriod;
                                        'Quantity'               = $data.extendedProperties.qty;
                                        'Savings Currency'       = $SavingsCurrency;
                                        'Annual Savings'         = "=$Savings";
                                        'Savings Region'         = $data.extendedProperties.region
                                    }
                                    $Adv += $obj
                                }
                            }
                        } catch {
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error processing Advisory item: ' + $_.Exception.Message)
                            continue
                        }
                    }
                    $script:Adv = $Adv
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Processed Advisory data inline (fallback): Count=' + $Adv.Count)
                }
            } else {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisories parameter is empty (Count=0)')
            }
        }
        
        # Now check if we have Advisory data (from any source: script scope, job, or direct processing)
        if ($null -ne $Adv) {
            # Ensure Adv is an array
            if ($Adv -isnot [System.Array]) {
                $Adv = @($Adv)
            }
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisory data found: Count=' + $Adv.Count)

            # Only generate sheet if we have Advisory data
            if ($Adv.Count -gt 0) {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Generating Advisor Sheet.')
                Build-ARIAdvisoryReport -File $File -Adv $Adv -TableStyle $TableStyle
                Write-Progress -Id 1 -activity 'Processing Advisories'  -Status "100% Complete." -Completed
            } else {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'No Advisory data to report - skipping Advisory sheet (Adv.Count = 0).')
            }

            Start-Sleep -Milliseconds 200
        } else {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'No Advisory data available - skipping Advisory sheet (Adv is null).')
        }
    } else {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'SkipAdvisory is set - skipping Advisory sheet generation.')
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