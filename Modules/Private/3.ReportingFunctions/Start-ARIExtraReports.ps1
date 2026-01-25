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
        # Check for Advisory job - it might be Completed, Failed, or not exist
        $AdvisoryJob = Get-Job -Name 'Advisory' -ErrorAction SilentlyContinue
        if ($null -ne $AdvisoryJob) {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisory job found: State=' + $AdvisoryJob.State)
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Receiving Advisory job results before Policy cleanup.')
            while (get-job -Name 'Advisory' | Where-Object { $_.State -eq 'Running' }) {
                Start-Sleep -Seconds 1
            }
            $script:Adv = Receive-Job -Name 'Advisory' -ErrorAction SilentlyContinue
            # Check for job errors
            if ($AdvisoryJob.State -eq 'Failed') {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisory job failed. Error: ' + ($AdvisoryJob.Error | Out-String))
            }
            Remove-Job -Name 'Advisory' -ErrorAction SilentlyContinue | Out-Null
            # Ensure Adv is an array for safe handling
            if ($null -eq $script:Adv) {
                $script:Adv = @()
            } elseif ($script:Adv -isnot [System.Array]) {
                $script:Adv = @($script:Adv)
            }
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisory data received and stored in script scope: Count=' + $script:Adv.Count)
        } else {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Advisory job not found when trying to receive results. Will use fallback processing.')
        }
    }

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Checking if Should Generate Policy Sheet.')
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'SkipPolicy parameter: IsPresent=' + $SkipPolicy.IsPresent + ', Value=' + $SkipPolicy)
    if (!$SkipPolicy.IsPresent) {
        # Check if Policy data is available from processed Policy.json cache (array format)
        # This happens when Policy.json contains processed records instead of raw PolicyAssign/PolicyDef/PolicySetDef
        $Pol = $null
        $policyCacheFile = $null
        
        # Try multiple paths to find Policy.json
        $possiblePaths = @()
        
        # Add Linux paths
        $possiblePaths += "/root/AzureResourceInventory/ReportCache/Policy.json"
        if ($null -ne $env:HOME) {
            $possiblePaths += (Join-Path $env:HOME "AzureResourceInventory/ReportCache/Policy.json")
        }
        
        # Add paths relative to Excel file (if File parameter is provided)
        if ($null -ne $File -and $File -ne '') {
            try {
                $fileParent = Split-Path $File -Parent -ErrorAction Stop
                if ($null -ne $fileParent -and $fileParent -ne '') {
                    $possiblePaths += (Join-Path $fileParent "ReportCache/Policy.json")
                    $possiblePaths += (Join-Path $fileParent "ReportCache\Policy.json")
                }
            } catch {
                # Ignore errors splitting file path
            }
        }
        
        # Add paths relative to current location
        try {
            $currentPath = (Get-Location).Path
            if ($null -ne $currentPath -and $currentPath -ne '') {
                $possiblePaths += (Join-Path $currentPath "ReportCache\Policy.json")
                $possiblePaths += (Join-Path $currentPath "ReportCache/Policy.json")
            }
        } catch {
            # Ignore errors getting current location
        }
        
        # Add relative paths
        $possiblePaths += ".\ReportCache\Policy.json"
        $possiblePaths += ".\ReportCache/Policy.json"
        
        foreach ($testPath in $possiblePaths) {
            if ($null -ne $testPath -and (Test-Path $testPath)) {
                $policyCacheFile = $testPath
                break
            }
        }
        
        # Also check ReportCache variable if available
        if ($null -eq $policyCacheFile) {
            try {
                $reportCacheVar = Get-Variable -Name 'ReportCache' -Scope 'Script' -ErrorAction SilentlyContinue
                if ($null -ne $reportCacheVar -and $null -ne $reportCacheVar.Value) {
                    $reportCachePath = $reportCacheVar.Value
                    $policyPath = Join-Path $reportCachePath "Policy.json"
                    if (Test-Path $policyPath) {
                        $policyCacheFile = $policyPath
                    }
                }
            } catch {
                # Ignore errors accessing ReportCache variable
            }
        }
        
        # Check for Policy job FIRST before checking Policy.json
        # This ensures we don't miss the Policy job if it's still running
        $policyJobExists = $false
        $policyJob = Get-Job -Name 'Policy' -ErrorAction SilentlyContinue
        if ($null -ne $policyJob) {
            $policyJobExists = $true
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Policy job found: State=' + $policyJob.State)
        }
        
        $hasRawPolicyData = $false
        if ($null -ne $policyCacheFile -and (Test-Path $policyCacheFile)) {
            try {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Checking Policy cache file: ' + $policyCacheFile)
                $policyCacheData = Get-Content $policyCacheFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                # Check if it's processed format (array) vs raw format (object with PolicyAssign/PolicyDef properties)
                if ($policyCacheData -is [System.Array]) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Found processed Policy data in cache (array format) - using directly')
                    $Pol = $policyCacheData
                    # Ensure Pol is an array
                    if ($null -eq $Pol) {
                        $Pol = @()
                    } elseif ($Pol -isnot [System.Array]) {
                        $Pol = @($Pol)
                    }
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Loaded ' + $Pol.Count + ' processed Policy record(s) from cache')
                } elseif ($policyCacheData.PSObject.Properties['PolicyAssign'] -or $policyCacheData.PSObject.Properties['PolicyDef']) {
                    # Raw Policy data found - Policy job should process it
                    $hasRawPolicyData = $true
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Found raw Policy data in cache (PolicyAssign/PolicyDef format) - Policy job will process it')
                    # Don't set $Pol here - let the Policy job process it
                    # Ensure we check for Policy job even if $Pol is null
                }
            } catch {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error loading Policy cache: ' + $_.Exception.Message)
            }
        }
        
        # If we have processed Policy data, use it directly (skip job)
        if ($null -ne $Pol -and $Pol.Count -gt 0) {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Generating Policy Sheet from processed cache data.')
            
            # Filter out records without human-readable names (GUIDs only)
            # Only include records where Policy has a displayName (not a GUID) and Initiative has a displayName (not "Policy Set: <guid>")
            $filteredPol = @()
            $excludedCount = 0
            foreach ($record in $Pol) {
                if ($null -eq $record) { continue }
                
                # Get Policy name and Initiative name
                $policyName = $null
                $initiativeName = $null
                
                if ($record -is [PSCustomObject]) {
                    if ($record.PSObject.Properties.Name -contains 'Policy') {
                        $policyName = $record.Policy
                    }
                    if ($record.PSObject.Properties.Name -contains 'Initiative') {
                        $initiativeName = $record.Initiative
                    }
                } elseif ($record -is [System.Collections.Hashtable] -or $record -is [System.Collections.IDictionary]) {
                    if ($record.ContainsKey('Policy')) {
                        $policyName = $record['Policy']
                    }
                    if ($record.ContainsKey('Initiative')) {
                        $initiativeName = $record['Initiative']
                    }
                }
                
                # Check if Policy name is a GUID (36-character GUID pattern)
                $isPolicyGuid = $false
                if ($null -ne $policyName -and $policyName -is [string]) {
                    # Check if it's a GUID pattern (36 chars: 8-4-4-4-12)
                    if ($policyName -match '^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$') {
                        $isPolicyGuid = $true
                    }
                }
                
                # Check if Initiative name indicates no displayName was found
                $isInitiativeGuid = $false
                if ($null -ne $initiativeName -and $initiativeName -is [string]) {
                    # If Initiative starts with "Policy Set:", it means no displayName was found
                    if ($initiativeName -match '^Policy Set:') {
                        $isInitiativeGuid = $true
                    }
                }
                
                # Only include records with human-readable names
                if (-not $isPolicyGuid -and -not $isInitiativeGuid) {
                    $filteredPol += $record
                } else {
                    $excludedCount++
                }
            }
            
            # Replace Pol with filtered results
            $Pol = $filteredPol
            
            if ($excludedCount -gt 0) {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Filtered out ' + $excludedCount + ' Policy record(s) without human-readable names (GUIDs only)')
            }
            
            # Aggressive memory cleanup before Policy sheet generation
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Running aggressive memory cleanup before Policy sheet Excel generation.')
            try {
                Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
                for ($i = 1; $i -le 5; $i++) {
                    [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $false)
                    [System.GC]::WaitForPendingFinalizers()
                    [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                }
                Clear-ARIMemory
            } catch {
                Write-Debug "  Warning: Pre-Policy memory cleanup had issues: $_"
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
            
            Write-Progress -Id 1 -activity 'Processing Policies' -Status "100% Complete." -Completed
        }
        # Otherwise, check for Policy job (for raw Policy data processing)
        # Use the $policyJobExists variable we checked earlier, or check again if job exists
        # Also check if we have raw Policy data - in that case, we MUST have a Policy job
        elseif ($hasRawPolicyData -or $policyJobExists -or (get-job -ErrorAction SilentlyContinue | Where-Object {$_.Name -eq 'Policy'}))
            {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Policy sheet generation path: hasRawPolicyData=' + $hasRawPolicyData + ', policyJobExists=' + $policyJobExists)
                
                # Check if Policy job actually exists before trying to receive it
                $policyJobToReceive = Get-Job -Name 'Policy' -ErrorAction SilentlyContinue
                if ($null -ne $policyJobToReceive) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Generating Policy Sheet from Policy job.')
                    
                    # Wait for Policy job to complete if it's still running
                    if ($policyJobToReceive.State -eq 'Running') {
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Policy job is still running - waiting for completion...')
                        while ($policyJobToReceive.State -eq 'Running') {
                            Write-Progress -Id 1 -activity 'Processing Policies' -Status "50% Complete." -PercentComplete 50
                            Start-Sleep -Seconds 2
                            $policyJobToReceive = Get-Job -Name 'Policy' -ErrorAction SilentlyContinue
                            if ($null -eq $policyJobToReceive) { break }
                        }
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Policy job completed: State=' + $policyJobToReceive.State)
                    }

                    # Aggressive memory cleanup BEFORE receiving Policy job results to free memory from previous operations
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Running aggressive memory cleanup before receiving Policy job results.')
                    try {
                        # Remove any other completed jobs first (Advisory was already received above)
                        # BUT preserve the Policy job until we receive its results
                        Get-Job -ErrorAction SilentlyContinue | Where-Object {$_.State -ne 'Running' -and $_.Name -ne 'Policy'} | Remove-Job -Force -ErrorAction SilentlyContinue
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

                    # Receive Policy job results
                    $Pol = Receive-Job -Name 'Policy' -ErrorAction SilentlyContinue
                    Remove-Job -Name 'Policy' -ErrorAction SilentlyContinue | Out-Null
                } else {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Policy job not found - will process raw Policy data directly from cache')
                    $Pol = $null
                }
                
                # If Policy job returned null and we have raw Policy data, process it directly
                if ($null -eq $Pol -and $hasRawPolicyData -and $null -ne $policyCacheFile -and (Test-Path $policyCacheFile)) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Policy job returned null - processing raw Policy data directly from cache')
                    try {
                        $policyCacheData = Get-Content $policyCacheFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                        if ($policyCacheData.PSObject.Properties['PolicyAssign'] -or $policyCacheData.PSObject.Properties['PolicyDef']) {
                            # Get Subscriptions from script scope or parameter
                            $SubsForPolicy = $null
                            try {
                                $subsVar = Get-Variable -Name 'Subscriptions' -Scope 'Script' -ErrorAction SilentlyContinue
                                if ($null -ne $subsVar) {
                                    $SubsForPolicy = $subsVar.Value
                                }
                            } catch {
                                # Try to get from function parameter
                                if ($null -ne $Subscriptions) {
                                    $SubsForPolicy = $Subscriptions
                                }
                            }
                            
                            # Ensure Subscriptions is an array
                            if ($null -eq $SubsForPolicy) {
                                $SubsForPolicy = @()
                            } elseif ($SubsForPolicy -isnot [System.Array]) {
                                $SubsForPolicy = @($SubsForPolicy)
                            }
                            
                            # Extract raw Policy data
                            $PolicyAssignRaw = if ($policyCacheData.PSObject.Properties['PolicyAssign']) { $policyCacheData.PolicyAssign } else { @{ policyAssignments = @() } }
                            $PolicyDefRaw = if ($policyCacheData.PSObject.Properties['PolicyDef']) { $policyCacheData.PolicyDef } else { @() }
                            $PolicySetDefRaw = if ($policyCacheData.PSObject.Properties['PolicySetDef']) { $policyCacheData.PolicySetDef } else { @() }
                            
                            # Ensure arrays
                            if ($null -eq $PolicyDefRaw) { $PolicyDefRaw = @() }
                            elseif ($PolicyDefRaw -isnot [System.Array]) { $PolicyDefRaw = @($PolicyDefRaw) }
                            if ($null -eq $PolicySetDefRaw) { $PolicySetDefRaw = @() }
                            elseif ($PolicySetDefRaw -isnot [System.Array]) { $PolicySetDefRaw = @($PolicySetDefRaw) }
                            
                            # Merge PolicyDef/PolicySetDef from PolicyBatch.json if it exists (batch data has better metadata)
                            # Collect from ALL batch PolicyBatch.json files and merged PolicyBatch.json
                            $reportCacheDir = Split-Path $policyCacheFile -Parent
                            $batchPolicyDefs = @()
                            $batchPolicySetDefs = @()
                            $batchesProcessed = 0
                            
                            # 1. Check merged PolicyBatch.json first (ReportCache_Merged)
                            $mergedCacheDir = Join-Path (Split-Path $reportCacheDir -Parent) "ReportCache_Merged"
                            $mergedPolicyBatchFile = if ($mergedCacheDir -and (Test-Path $mergedCacheDir)) { Join-Path $mergedCacheDir "PolicyBatch.json" } else { $null }
                            if ($null -ne $mergedPolicyBatchFile -and (Test-Path $mergedPolicyBatchFile)) {
                                try {
                                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Found merged PolicyBatch.json - loading PolicyDef/PolicySetDef')
                                    $mergedBatchData = Get-Content $mergedPolicyBatchFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                                    if ($mergedBatchData -is [PSCustomObject] -or $mergedBatchData -is [System.Collections.Hashtable]) {
                                        if ($mergedBatchData.PSObject.Properties.Name -contains 'PolicyDef' -or $mergedBatchData.ContainsKey('PolicyDef')) {
                                            $defs = if ($mergedBatchData -is [PSCustomObject]) { $mergedBatchData.PolicyDef } else { $mergedBatchData['PolicyDef'] }
                                            if ($defs -is [System.Array]) { $batchPolicyDefs += $defs } elseif ($null -ne $defs) { $batchPolicyDefs += @($defs) }
                                        }
                                        if ($mergedBatchData.PSObject.Properties.Name -contains 'PolicySetDef' -or $mergedBatchData.ContainsKey('PolicySetDef')) {
                                            $setDefs = if ($mergedBatchData -is [PSCustomObject]) { $mergedBatchData.PolicySetDef } else { $mergedBatchData['PolicySetDef'] }
                                            if ($setDefs -is [System.Array]) { $batchPolicySetDefs += $setDefs } elseif ($null -ne $setDefs) { $batchPolicySetDefs += @($setDefs) }
                                        }
                                        $batchesProcessed++
                                    }
                                } catch {
                                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Warning: Failed to load merged PolicyBatch.json: ' + $_.Exception.Message)
                                }
                            }
                            
                            # 2. Check batch directories (ReportCache_Batch*) - collect from ALL batches
                            $batchDirs = Get-ChildItem -Path (Split-Path $reportCacheDir -Parent) -Directory -Filter "ReportCache_Batch*" -ErrorAction SilentlyContinue
                            foreach ($batchDir in $batchDirs) {
                                $batchPolicyBatchFile = Join-Path $batchDir.FullName "PolicyBatch.json"
                                if (Test-Path $batchPolicyBatchFile) {
                                    try {
                                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Found PolicyBatch.json in ' + $batchDir.Name + ' - loading PolicyDef/PolicySetDef')
                                        $batchData = Get-Content $batchPolicyBatchFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                                        if ($batchData -is [PSCustomObject] -or $batchData -is [System.Collections.Hashtable]) {
                                            if ($batchData.PSObject.Properties.Name -contains 'PolicyDef' -or $batchData.ContainsKey('PolicyDef')) {
                                                $defs = if ($batchData -is [PSCustomObject]) { $batchData.PolicyDef } else { $batchData['PolicyDef'] }
                                                if ($defs -is [System.Array]) { $batchPolicyDefs += $defs } elseif ($null -ne $defs) { $batchPolicyDefs += @($defs) }
                                            }
                                            if ($batchData.PSObject.Properties.Name -contains 'PolicySetDef' -or $batchData.ContainsKey('PolicySetDef')) {
                                                $setDefs = if ($batchData -is [PSCustomObject]) { $batchData.PolicySetDef } else { $batchData['PolicySetDef'] }
                                                if ($setDefs -is [System.Array]) { $batchPolicySetDefs += $setDefs } elseif ($null -ne $setDefs) { $batchPolicySetDefs += @($setDefs) }
                                            }
                                            $batchesProcessed++
                                        }
                                    } catch {
                                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Warning: Failed to load PolicyBatch.json from ' + $batchDir.Name + ': ' + $_.Exception.Message)
                                    }
                                }
                            }
                            
                            # 3. Check same directory as Policy.json (fallback)
                            $sameDirPolicyBatchFile = Join-Path $reportCacheDir "PolicyBatch.json"
                            if ($batchesProcessed -eq 0 -and (Test-Path $sameDirPolicyBatchFile)) {
                                try {
                                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Found PolicyBatch.json in same directory as Policy.json - loading PolicyDef/PolicySetDef')
                                    $sameDirBatchData = Get-Content $sameDirPolicyBatchFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                                    if ($sameDirBatchData -is [PSCustomObject] -or $sameDirBatchData -is [System.Collections.Hashtable]) {
                                        if ($sameDirBatchData.PSObject.Properties.Name -contains 'PolicyDef' -or $sameDirBatchData.ContainsKey('PolicyDef')) {
                                            $defs = if ($sameDirBatchData -is [PSCustomObject]) { $sameDirBatchData.PolicyDef } else { $sameDirBatchData['PolicyDef'] }
                                            if ($defs -is [System.Array]) { $batchPolicyDefs += $defs } elseif ($null -ne $defs) { $batchPolicyDefs += @($defs) }
                                        }
                                        if ($sameDirBatchData.PSObject.Properties.Name -contains 'PolicySetDef' -or $sameDirBatchData.ContainsKey('PolicySetDef')) {
                                            $setDefs = if ($sameDirBatchData -is [PSCustomObject]) { $sameDirBatchData.PolicySetDef } else { $sameDirBatchData['PolicySetDef'] }
                                            if ($setDefs -is [System.Array]) { $batchPolicySetDefs += $setDefs } elseif ($null -ne $setDefs) { $batchPolicySetDefs += @($setDefs) }
                                        }
                                        $batchesProcessed++
                                    }
                                } catch {
                                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Warning: Failed to load PolicyBatch.json from same directory: ' + $_.Exception.Message)
                                }
                            }
                            
                            # Merge: Add batch data first (takes precedence), then early collection data, then deduplicate by ID
                            if ($batchesProcessed -gt 0) {
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Found PolicyBatch.json files - merging PolicyDef/PolicySetDef with Policy.json data (processed ' + $batchesProcessed + ' batch file(s))')
                                
                                # Deduplicate batch PolicyDefs/PolicySetDefs by ID first
                                $batchPolicyDefs = $batchPolicyDefs | Sort-Object -Property id -Unique
                                $batchPolicySetDefs = $batchPolicySetDefs | Sort-Object -Property id -Unique
                                
                                # Merge: Add batch data first (takes precedence), then early collection data, then deduplicate by ID
                                $allPolicyDefs = @()
                                $allPolicyDefs += $batchPolicyDefs  # Batch data first (preferred)
                                $allPolicyDefs += $PolicyDefRaw     # Then early collection data
                                $PolicyDefRaw = $allPolicyDefs | Sort-Object -Property id -Unique
                                
                                $allPolicySetDefs = @()
                                $allPolicySetDefs += $batchPolicySetDefs  # Batch data first (preferred)
                                $allPolicySetDefs += $PolicySetDefRaw     # Then early collection data
                                $PolicySetDefRaw = $allPolicySetDefs | Sort-Object -Property id -Unique
                                
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Merged PolicyDef/PolicySetDef: PolicyDef=' + $PolicyDefRaw.Count + ' (batch: ' + $batchPolicyDefs.Count + ' unique), PolicySetDef=' + $PolicySetDefRaw.Count + ' (batch: ' + $batchPolicySetDefs.Count + ' unique)')
                            }
                            
                            # Handle PolicyAssign structure - it may be an array or an object with policyAssignments property
                            if ($null -ne $PolicyAssignRaw) {
                                if ($PolicyAssignRaw -is [System.Array]) {
                                    # Direct array - wrap in hashtable with policyAssignments property
                                    $PolicyAssignRaw = @{ policyAssignments = $PolicyAssignRaw }
                                } elseif ($PolicyAssignRaw -is [PSCustomObject] -or $PolicyAssignRaw -is [System.Collections.Hashtable]) {
                                    # Already has structure, check for policyAssignments property
                                    if (-not ($PolicyAssignRaw.policyAssignments -or $PolicyAssignRaw.ContainsKey('policyAssignments'))) {
                                        # Convert to hashtable with policyAssignments property
                                        $PolicyAssignRaw = @{ policyAssignments = @() }
                                    }
                                } else {
                                    # Single value - wrap in hashtable
                                    $PolicyAssignRaw = @{ policyAssignments = @($PolicyAssignRaw) }
                                }
                            } else {
                                $PolicyAssignRaw = @{ policyAssignments = @() }
                            }
                            
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Calling Start-ARIPolicyJob directly with ' + $SubsForPolicy.Count + ' subscription(s)')
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'PolicyAssign structure: ' + ($PolicyAssignRaw.GetType().Name) + ', PolicyDef count: ' + $PolicyDefRaw.Count + ', PolicySetDef count: ' + $PolicySetDefRaw.Count)
                            try {
                                $Pol = Start-ARIPolicyJob -Subscriptions $SubsForPolicy -PolicySetDef $PolicySetDefRaw -PolicyAssign $PolicyAssignRaw -PolicyDef $PolicyDefRaw
                                if ($null -eq $Pol) {
                                    $Pol = @()
                                } elseif ($Pol -isnot [System.Array]) {
                                    $Pol = @($Pol)
                                }
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob returned ' + $Pol.Count + ' Policy record(s)')
                            } catch {
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error calling Start-ARIPolicyJob: ' + $_.Exception.Message)
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Stack trace: ' + $_.ScriptStackTrace)
                                $Pol = @()
                            }
                        }
                    } catch {
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error processing raw Policy data directly: ' + $_.Exception.Message)
                        $Pol = @()
                    }
                }
                
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Received Policy job results: Pol is null=' + ($null -eq $Pol))
                if ($null -ne $Pol) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Pol type=' + ($Pol.GetType().Name))
                }
                
                # Ensure Pol is an array for safe handling
                if ($null -eq $Pol) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'WARNING: Policy job returned null results')
                    $Pol = @()
                } elseif ($Pol -isnot [System.Array]) {
                    $Pol = @($Pol)
                }
                
                # Filter out records without human-readable names (GUIDs only)
                # Only include records where Policy has a displayName (not a GUID) and Initiative has a displayName (not "Policy Set: <guid>")
                $filteredPol = @()
                $excludedCount = 0
                foreach ($record in $Pol) {
                    if ($null -eq $record) { continue }
                    
                    # Get Policy name and Initiative name
                    $policyName = $null
                    $initiativeName = $null
                    
                    if ($record -is [PSCustomObject]) {
                        if ($record.PSObject.Properties.Name -contains 'Policy') {
                            $policyName = $record.Policy
                        }
                        if ($record.PSObject.Properties.Name -contains 'Initiative') {
                            $initiativeName = $record.Initiative
                        }
                    } elseif ($record -is [System.Collections.Hashtable] -or $record -is [System.Collections.IDictionary]) {
                        if ($record.ContainsKey('Policy')) {
                            $policyName = $record['Policy']
                        }
                        if ($record.ContainsKey('Initiative')) {
                            $initiativeName = $record['Initiative']
                        }
                    }
                    
                    # Check if Policy name is a GUID (36-character GUID pattern)
                    $isPolicyGuid = $false
                    if ($null -ne $policyName -and $policyName -is [string]) {
                        # Check if it's a GUID pattern (36 chars: 8-4-4-4-12)
                        if ($policyName -match '^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$') {
                            $isPolicyGuid = $true
                        }
                    }
                    
                    # Check if Initiative name indicates no displayName was found
                    $isInitiativeGuid = $false
                    if ($null -ne $initiativeName -and $initiativeName -is [string]) {
                        # If Initiative starts with "Policy Set:", it means no displayName was found
                        if ($initiativeName -match '^Policy Set:') {
                            $isInitiativeGuid = $true
                        }
                    }
                    
                    # Only include records with human-readable names
                    if (-not $isPolicyGuid -and -not $isInitiativeGuid) {
                        $filteredPol += $record
                    } else {
                        $excludedCount++
                    }
                }
                
                # Replace Pol with filtered results
                $Pol = $filteredPol
                
                if ($excludedCount -gt 0) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Filtered out ' + $excludedCount + ' Policy record(s) without human-readable names (GUIDs only)')
                }
                
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Policy data ready for sheet generation: Count=' + $Pol.Count)
                
                # If Policy job returned empty results, log warning but still try to generate sheet
                if ($Pol.Count -eq 0) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'WARNING: Policy job returned empty results - Policy sheet will be empty')
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
                                        # Handle advisories WITH resourceId (resource-level recommendations)
                                        if ($null -ne $data -and $null -ne $data.resourceMetadata -and $null -ne $data.resourceMetadata.resourceId) {
                                            $Savings = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){$data.extendedProperties.annualSavingsAmount}Else{0}
                                            $SavingsCurrency = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){$data.extendedProperties.savingsCurrency}Else{'USD'}
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
                                                'SKU'                    = if ($null -ne $data.extendedProperties) { $data.extendedProperties.sku } else { '' };
                                                'Term'                   = if ($null -ne $data.extendedProperties) { $data.extendedProperties.term } else { '' };
                                                'Look-back Period'       = if ($null -ne $data.extendedProperties) { $data.extendedProperties.lookbackPeriod } else { '' };
                                                'Quantity'               = if ($null -ne $data.extendedProperties) { $data.extendedProperties.qty } else { '' };
                                                'Savings Currency'       = $SavingsCurrency;
                                                'Annual Savings'         = "=$Savings";
                                                'Savings Region'         = if ($null -ne $data.extendedProperties) { $data.extendedProperties.region } else { '' }
                                            }
                                            $Adv += $obj
                                        }
                                        # Handle advisories WITHOUT resourceId (subscription-level or management group-level recommendations)
                                        elseif ($null -ne $data) {
                                            # Extract subscription ID from advisory ID if available
                                            $Subscription = ''
                                            if ($null -ne $advItem.id) {
                                                # Advisory ID format: /subscriptions/{subId}/providers/Microsoft.Advisor/recommendations/{recId}
                                                $idParts = $advItem.id -split '/'
                                                $subIndex = [array]::IndexOf($idParts, 'subscriptions')
                                                if ($subIndex -ge 0 -and $subIndex + 1 -lt $idParts.Count) {
                                                    $Subscription = $idParts[$subIndex + 1]
                                                }
                                            }
                                            
                                            # Use impactedField/impactedValue for resource type/name
                                            $ResourceType = if ($null -ne $data.impactedField) { $data.impactedField } else { '' }
                                            $ResourceName = if ($null -ne $data.impactedValue) { $data.impactedValue } else { '' }
                                            
                                            $Savings = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){$data.extendedProperties.annualSavingsAmount}Else{0}
                                            $SavingsCurrency = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){$data.extendedProperties.savingsCurrency}Else{'USD'}
                                            
                                            $obj = @{
                                                'Subscription'           = $Subscription;
                                                'Resource Group'         = '';
                                                'Resource Type'          = $ResourceType;
                                                'Name'                   = $ResourceName;
                                                'Detailed Type'          = '';
                                                'Detailed Name'          = '';
                                                'Category'               = if ($null -ne $data.category) { $data.category } else { '' };
                                                'Impact'                 = if ($null -ne $data.impact) { $data.impact } else { '' };
                                                'Description'            = if ($null -ne $data.shortDescription) { $data.shortDescription.problem } else { '' };
                                                'SKU'                    = if ($null -ne $data.extendedProperties) { $data.extendedProperties.sku } else { '' };
                                                'Term'                   = if ($null -ne $data.extendedProperties) { $data.extendedProperties.term } else { '' };
                                                'Look-back Period'       = if ($null -ne $data.extendedProperties) { $data.extendedProperties.lookbackPeriod } else { '' };
                                                'Quantity'               = if ($null -ne $data.extendedProperties) { $data.extendedProperties.qty } else { '' };
                                                'Savings Currency'       = $SavingsCurrency;
                                                'Annual Savings'         = "=$Savings";
                                                'Savings Region'         = if ($null -ne $data.extendedProperties) { $data.extendedProperties.region } else { '' }
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
                                    # Handle advisories WITH resourceId (resource-level recommendations)
                                    if ($null -ne $data -and $null -ne $data.resourceMetadata -and $null -ne $data.resourceMetadata.resourceId) {
                                        $Savings = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){$data.extendedProperties.annualSavingsAmount}Else{0}
                                        $SavingsCurrency = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){$data.extendedProperties.savingsCurrency}Else{'USD'}
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
                                    # Handle advisories WITHOUT resourceId (subscription-level or management group-level recommendations)
                                    elseif ($null -ne $data) {
                                        # Extract subscription ID from advisory ID if available
                                        $Subscription = ''
                                        if ($null -ne $advItem.id) {
                                            # Advisory ID format: /subscriptions/{subId}/providers/Microsoft.Advisor/recommendations/{recId}
                                            $idParts = $advItem.id -split '/'
                                            $subIndex = [array]::IndexOf($idParts, 'subscriptions')
                                            if ($subIndex -ge 0 -and $subIndex + 1 -lt $idParts.Count) {
                                                $Subscription = $idParts[$subIndex + 1]
                                            }
                                        }
                                        
                                        # Use impactedField/impactedValue for resource type/name
                                        $ResourceType = if ($null -ne $data.impactedField) { $data.impactedField } else { '' }
                                        $ResourceName = if ($null -ne $data.impactedValue) { $data.impactedValue } else { '' }
                                        
                                        $Savings = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){$data.extendedProperties.annualSavingsAmount}Else{0}
                                        $SavingsCurrency = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){$data.extendedProperties.savingsCurrency}Else{'USD'}
                                        
                                        $obj = @{
                                            'Subscription'           = $Subscription;
                                            'Resource Group'         = '';
                                            'Resource Type'          = $ResourceType;
                                            'Name'                   = $ResourceName;
                                            'Detailed Type'          = '';
                                            'Detailed Name'          = '';
                                            'Category'               = if ($null -ne $data.category) { $data.category } else { '' };
                                            'Impact'                 = if ($null -ne $data.impact) { $data.impact } else { '' };
                                            'Description'            = if ($null -ne $data.shortDescription) { $data.shortDescription.problem } else { '' };
                                            'SKU'                    = if ($null -ne $data.extendedProperties) { $data.extendedProperties.sku } else { '' };
                                            'Term'                   = if ($null -ne $data.extendedProperties) { $data.extendedProperties.term } else { '' };
                                            'Look-back Period'       = if ($null -ne $data.extendedProperties) { $data.extendedProperties.lookbackPeriod } else { '' };
                                            'Quantity'               = if ($null -ne $data.extendedProperties) { $data.extendedProperties.qty } else { '' };
                                            'Savings Currency'       = $SavingsCurrency;
                                            'Annual Savings'         = "=$Savings";
                                            'Savings Region'         = if ($null -ne $data.extendedProperties) { $data.extendedProperties.region } else { '' }
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
                                # Handle advisories WITH resourceId (resource-level recommendations)
                                if ($null -ne $data -and $null -ne $data.resourceMetadata -and $null -ne $data.resourceMetadata.resourceId) {
                                    $Savings = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){$data.extendedProperties.annualSavingsAmount}Else{0}
                                    $SavingsCurrency = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){$data.extendedProperties.savingsCurrency}Else{'USD'}
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
                                # Handle advisories WITHOUT resourceId (subscription-level or management group-level recommendations)
                                elseif ($null -ne $data) {
                                    # Extract subscription ID from advisory ID if available
                                    $Subscription = ''
                                    if ($null -ne $advItem.id) {
                                        # Advisory ID format: /subscriptions/{subId}/providers/Microsoft.Advisor/recommendations/{recId}
                                        $idParts = $advItem.id -split '/'
                                        $subIndex = [array]::IndexOf($idParts, 'subscriptions')
                                        if ($subIndex -ge 0 -and $subIndex + 1 -lt $idParts.Count) {
                                            $Subscription = $idParts[$subIndex + 1]
                                        }
                                    }
                                    
                                    # Use impactedField/impactedValue for resource type/name
                                    $ResourceType = if ($null -ne $data.impactedField) { $data.impactedField } else { '' }
                                    $ResourceName = if ($null -ne $data.impactedValue) { $data.impactedValue } else { '' }
                                    
                                    $Savings = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){$data.extendedProperties.annualSavingsAmount}Else{0}
                                    $SavingsCurrency = if ($null -ne $data.extendedProperties -and -not [string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){$data.extendedProperties.savingsCurrency}Else{'USD'}
                                    
                                    $obj = @{
                                        'Subscription'           = $Subscription;
                                        'Resource Group'         = '';
                                        'Resource Type'          = $ResourceType;
                                        'Name'                   = $ResourceName;
                                        'Detailed Type'          = '';
                                        'Detailed Name'          = '';
                                        'Category'               = if ($null -ne $data.category) { $data.category } else { '' };
                                        'Impact'                 = if ($null -ne $data.impact) { $data.impact } else { '' };
                                        'Description'            = if ($null -ne $data.shortDescription) { $data.shortDescription.problem } else { '' };
                                        'SKU'                    = if ($null -ne $data.extendedProperties) { $data.extendedProperties.sku } else { '' };
                                        'Term'                   = if ($null -ne $data.extendedProperties) { $data.extendedProperties.term } else { '' };
                                        'Look-back Period'       = if ($null -ne $data.extendedProperties) { $data.extendedProperties.lookbackPeriod } else { '' };
                                        'Quantity'               = if ($null -ne $data.extendedProperties) { $data.extendedProperties.qty } else { '' };
                                        'Savings Currency'       = $SavingsCurrency;
                                        'Annual Savings'         = "=$Savings";
                                        'Savings Region'         = if ($null -ne $data.extendedProperties) { $data.extendedProperties.region } else { '' }
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

    # Try multiple sources for Subscriptions data:
    # 1. Script-scoped variable (from job results received earlier before cleanup)
    # 2. Subscriptions job (if still available)
    $AzSubs = $script:AzSubs
    
    if ($null -eq $AzSubs -or ($AzSubs -is [System.Array] -and $AzSubs.Count -eq 0)) {
        # Fallback: try to receive from job if script-scoped variable is null or empty
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Script-scoped Subscriptions data is null or empty - checking for Subscriptions job...')
        $SubscriptionsJob = Get-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
        if ($null -ne $SubscriptionsJob) {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job found: State=' + $SubscriptionsJob.State)
            # Wait for job to complete if still running
            while ($SubscriptionsJob.State -eq 'Running') {
                Write-Progress -Id 1 -activity 'Processing Subscriptions' -Status "50% Complete." -PercentComplete 50
                Start-Sleep -Seconds 2
                $SubscriptionsJob = Get-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
                if ($null -eq $SubscriptionsJob) { break }
            }
            if ($null -ne $SubscriptionsJob) {
                # Check if job failed
                if ($SubscriptionsJob.State -eq 'Failed') {
                    $jobError = Receive-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job failed: ' + ($jobError | Out-String))
                    $AzSubs = @()
                } else {
                    try {
                        $jobOutput = Receive-Job -Name 'Subscriptions' -ErrorAction Stop
                        
                        # Less aggressive filtering: Only filter out DebugRecord objects
                        if ($jobOutput -is [System.Array]) {
                            # Filter out DebugRecord objects, keep everything else
                            $dataObjects = $jobOutput | Where-Object { $_ -isnot [System.Management.Automation.DebugRecord] }
                            # Prefer PSCustomObjects (the actual data)
                            $psCustomObjects = $dataObjects | Where-Object { $_ -is [PSCustomObject] }
                            if ($psCustomObjects.Count -gt 0) {
                                $AzSubs = $psCustomObjects
                            } elseif ($dataObjects.Count -gt 0) {
                                # Use other objects if no PSCustomObjects found (might be deserialized)
                                $AzSubs = $dataObjects
                            } else {
                                $AzSubs = @()
                            }
                        } elseif ($jobOutput -is [PSCustomObject]) {
                            $AzSubs = $jobOutput
                        } elseif ($jobOutput -isnot [System.Management.Automation.DebugRecord] -and $jobOutput -isnot [string]) {
                            # Might be a deserialized object - try to use it
                            $AzSubs = $jobOutput
                        } else {
                            $AzSubs = @()
                        }
                        
                        # Store in script scope for safety
                        $script:AzSubs = $AzSubs
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Received Subscriptions data from job: Count=' + (if ($AzSubs -is [System.Array]) { $AzSubs.Count } else { if ($null -eq $AzSubs) { 0 } else { 1 } }))
                    } catch {
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error receiving Subscriptions job results: ' + $_.Exception.Message)
                        $AzSubs = @()
                    }
                }
                Remove-Job -Name 'Subscriptions' -ErrorAction SilentlyContinue | Out-Null
            }
        } else {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Subscriptions job not found - may have completed and been removed already.')
        }
    } else {
        $subsCount = if ($AzSubs -is [System.Array]) { $AzSubs.Count } elseif ($null -eq $AzSubs) { 0 } else { 1 }
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Using script-scoped Subscriptions data: Count=' + $subsCount)
    }
    
    # Ensure AzSubs is an array for safe handling
    if ($null -eq $AzSubs) {
        $AzSubs = @()
    } elseif ($AzSubs -isnot [System.Array]) {
        # If it's a single object, wrap it in an array
        $AzSubs = @($AzSubs)
    }
    
    # If still empty, log a warning
    if ($AzSubs.Count -eq 0) {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Warning: No Subscriptions data available (Count=0) - Subscriptions sheet will be empty or skipped')
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'No subscription data to report - skipping Subscriptions sheet')
    } else {
        try {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Calling Build-ARISubsReport with ' + $AzSubs.Count + ' subscription record(s)')
            Build-ARISubsReport -File $File -Sub $AzSubs -IncludeCosts $IncludeCosts -TableStyle $TableStyle
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Build-ARISubsReport completed successfully')
        } catch {
            # Safe error handling - check property existence before accessing
            $errorMsg = if ($null -ne $_ -and $null -ne $_.Exception) { $_.Exception.Message } else { "Unknown error" }
            
            $errorLine = "Unknown"
            $errorFunc = "Unknown"
            
            try {
                if ($null -ne $_ -and $null -ne $_.InvocationInfo) {
                    $errorLine = if ($null -ne $_.InvocationInfo.ScriptLineNumber) { $_.InvocationInfo.ScriptLineNumber } else { "Unknown" }
                    # Check if FunctionName property exists before accessing
                    if ($_.InvocationInfo.PSObject.Properties.Name -contains 'FunctionName') {
                        $errorFunc = $_.InvocationInfo.FunctionName
                    }
                }
            } catch {
                # Ignore errors accessing InvocationInfo
            }
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error in Build-ARISubsReport: ' + $errorMsg)
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error at line: ' + $errorLine + ', Function: ' + $errorFunc)
            Write-Host "  [ERROR] Failed to generate Subscriptions sheet: $errorMsg" -ForegroundColor Red
            # Don't throw - continue with other sheets
        }
    }

    Clear-ARIMemory

    Write-Progress -activity 'Azure Resource Inventory Subscriptions' -Status "100% Complete." -Completed

    Write-Progress -activity 'Azure Inventory' -Status "80% Complete." -PercentComplete 80 -CurrentOperation "Completed Extra Resources Reporting.."
}