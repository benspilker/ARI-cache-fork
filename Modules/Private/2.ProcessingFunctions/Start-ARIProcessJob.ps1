<#
.Synopsis
Module responsible for starting the processing jobs for Azure Resources.

.DESCRIPTION
This module creates and manages jobs to process Azure Resources in batches based on the environment size. It ensures efficient resource processing and avoids CPU overload.

.Link
https://github.com/microsoft/ARI/Modules/Private/2.ProcessingFunctions/Start-ARIProcessJob.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI).

.NOTES
Version: 3.6.5
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Start-ARIProcessJob {
    Param($Resources, $Retirements, $Subscriptions, $DefaultPath, $Heavy, $InTag, $Unsupported)

    function Write-ARIModuleMemoryCheckpoint {
        param(
            [string]$Stage,
            [string]$ModuleName = ''
        )
        try {
            $proc = Get-Process -Id $PID -ErrorAction Stop
            $wsMb = [math]::Round(($proc.WorkingSet64 / 1MB), 2)
            $privateMb = [math]::Round(($proc.PrivateMemorySize64 / 1MB), 2)
            $virtualMb = [math]::Round(($proc.VirtualMemorySize64 / 1MB), 2)
            $moduleSuffix = if ([string]::IsNullOrWhiteSpace($ModuleName)) { '' } else { " module=$ModuleName" }
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss') + ' - ' + "[MEMCHK] stage=$Stage$moduleSuffix ws_mb=$wsMb private_mb=$privateMb virtual_mb=$virtualMb")
        } catch {
            $moduleSuffix = if ([string]::IsNullOrWhiteSpace($ModuleName)) { '' } else { " module=$ModuleName" }
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss') + ' - ' + "[MEMCHK] stage=$Stage$moduleSuffix failed=$($_.Exception.Message)")
        }
    }

    Write-Progress -activity 'Azure Inventory' -Status "22% Complete." -PercentComplete 22 -CurrentOperation "Creating Jobs to Process Data.."

    # PATCHED: Limit concurrent jobs for Windmill memory constraints
    # Use 1 concurrent job when resource count >= 1000, else 2
    $resourceCount = 0
    if ($null -ne $Resources) {
        if ($Resources -is [System.Array]) {
            $resourceCount = $Resources.Count
        } elseif ($Resources -is [System.Collections.ICollection]) {
            $resourceCount = $Resources.Count
        } else {
            $resourceCount = 1
        }
    }
    $EnvSizeLooper = if ($resourceCount -ge 1000) { 1 } else { 2 }
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'[PATCHED] Limiting concurrent jobs to '+$EnvSizeLooper+' for Windmill memory constraints (resource count: '+$resourceCount+')')
    
    # Original logic (commented out for reference):
    # switch ($Resources.count)
    # {
    #     {$_ -le 12500}
    #         {
    #             Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Regular Size Environment. Jobs will be run in parallel.')
    #             $EnvSizeLooper = 20
    #         }
    #     {$_ -gt 12500 -and $_ -le 50000}
    #         {
    #             Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Medium Size Environment. Jobs will be run in batches of 8.')
    #             $EnvSizeLooper = 8
    #         }
    #     {$_ -gt 50000}
    #         {
    #             Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Large Environment Detected.')
    #             $EnvSizeLooper = 5
    #             Write-Host ('Jobs will be run in small batches to avoid CPU and Memory Overload.') -ForegroundColor Red
    #         }
    # }
    #
    # if ($Heavy.IsPresent -or $InTag.IsPresent)
    #     {
    #         Write-Host ('Heavy Mode or InTag Mode Detected. Jobs will be run in small batches to avoid CPU and Memory Overload.') -ForegroundColor Red
    #         $EnvSizeLooper = 5
    #     }

    $ParentPath = (get-item $PSScriptRoot).parent.parent
    $InventoryModulesPath = Join-Path $ParentPath 'Public' 'InventoryModules'
    $ModuleFolders = Get-ChildItem -Path $InventoryModulesPath -Directory

    $JobLoop = 1
    $TotalFolders = $ModuleFolders.count

    # PATCHED: Stream JSON to disk via StreamWriter to bound parent peak memory.
    # Replaces the four runtime OOM patches (LightPayload + LightPayloadChunked +
    # StreamingJson + ChunkedStreamingJson). Each resource is serialized
    # individually via ConvertTo-Json, written to the StreamWriter, then the
    # heavy object slot in $Resources is nulled so the GC can reclaim it.
    #
    # WHY THIS WORKS (and the old light-payload patch did not):
    #   The light-payload build kept a LIVE REFERENCE to the heavy object's
    #   `properties` dictionary via `$r.properties`, so nulling $Resources[$i]
    #   did NOT free the heavy data - the light wrapper still pointed at it.
    #   This version writes the heavy object directly and drops the reference,
    #   so each element is genuinely collectable after it hits disk.
    #
    # PEAK MEMORY during the streaming loop:
    #   - 1 chunk's worth of $Resources still alive (~250 elements * ~30 KB = ~7.5 MB)
    #   - 1 transient JSON string per ConvertTo-Json call (~50 KB, freed each iter)
    #   - StreamWriter internal buffer (~4 KB)
    #   - Unrelated parent overhead (runspaces, jobs, modules) (~100-200 MB)
    # Total parent peak drops from ~1.6 GB to ~250 MB on 9,400-resource tenants.
    #
    # NOTE: This fixes the PARENT OOM. Children still materialize the full JSON
    # via ConvertFrom-Json, which can OOM the child runspace for very large
    # tenants. That's Path B (per-module JSON files) - tracked separately.
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'[STREAMING-JSON] Writing JSON to disk via StreamWriter (chunked, nulled per element)')
    $ResourcesJsonPath = Join-Path $DefaultPath ("ari-resources-" + [guid]::NewGuid().ToString("N") + ".json")
    $__chunkSize = 250
    $__rCount = 0
    if ($null -ne $Resources) {
        if ($Resources -is [System.Collections.ICollection]) {
            $__rCount = $Resources.Count
        } else {
            # Defensive: single-object case. Wrap so the streaming loop can index it.
            $Resources = @($Resources)
            $__rCount = 1
        }
    }
    # ARI_AI_RESOURCE_SCOPE_V2: build the AI subset in the same pass as the full
    # resource payload. Avoid reloading the full JSON file into memory just to
    # filter AI resources (the prior path used Get-Content -Raw | ConvertFrom-Json).
    $aiResourceTypePrefixes = @(
        'microsoft.cognitiveservices/accounts',
        'microsoft.machinelearningservices/workspaces',
        'microsoft.search/searchservices'
    )
    $aiProviderPrefixes = @(
        'microsoft.cognitiveservices/',
        'microsoft.machinelearningservices/',
        'microsoft.search/'
    )
    $aiDirectPayloadPath = Join-Path $DefaultPath ("ari-resources-ai-direct-" + [guid]::NewGuid().ToString("N") + ".json")
    $aiProviderPayloadPath = Join-Path $DefaultPath ("ari-resources-ai-provider-" + [guid]::NewGuid().ToString("N") + ".json")
    $aiDirectResourceCount = 0
    $aiProviderResourceCount = 0
    $__aiDirectFirst = $true
    $__aiProviderFirst = $true

    $__writer = $null
    $__aiDirectWriter = $null
    $__aiProviderWriter = $null
    try {
        $__writer = [System.IO.StreamWriter]::new($ResourcesJsonPath, $false, [System.Text.UTF8Encoding]::new($false))
        $__aiDirectWriter = [System.IO.StreamWriter]::new($aiDirectPayloadPath, $false, [System.Text.UTF8Encoding]::new($false))
        $__aiProviderWriter = [System.IO.StreamWriter]::new($aiProviderPayloadPath, $false, [System.Text.UTF8Encoding]::new($false))
        $__writer.Write('[')
        $__aiDirectWriter.Write('[')
        $__aiProviderWriter.Write('[')
        $__first = $true
        if ($__rCount -gt 0) {
            $__chunkCount = [math]::Ceiling($__rCount / $__chunkSize)
            for ($__ci = 0; $__ci -lt $__chunkCount; $__ci++) {
                $__start = $__ci * $__chunkSize
                $__end = [math]::Min($__start + $__chunkSize, $__rCount)
                for ($__ri = $__start; $__ri -lt $__end; $__ri++) {
                    $__r = $Resources[$__ri]
                    if ($null -eq $__r) { continue }
                    $__json = ConvertTo-Json -InputObject $__r -Depth 40 -Compress
                    if (-not $__first) { $__writer.Write(',') }
                    $__writer.Write($__json)
                    $__first = $false

                    $__resourceType = ''
                    try {
                        if ($__r.PSObject.Properties['type']) { $__resourceType = [string]$__r.type }
                        elseif ($__r.PSObject.Properties['TYPE']) { $__resourceType = [string]$__r.TYPE }
                        elseif ($__r.PSObject.Properties['resourceType']) { $__resourceType = [string]$__r.resourceType }
                        elseif ($__r.PSObject.Properties['ResourceType']) { $__resourceType = [string]$__r.ResourceType }
                    } catch {}
                    $__normalizedResourceType = $__resourceType.Trim().ToLowerInvariant()
                    $__isDirectAi = $false
                    $__isProviderAi = $false
                    if (-not [string]::IsNullOrWhiteSpace($__normalizedResourceType)) {
                        foreach ($__prefix in $aiResourceTypePrefixes) {
                            if ($__normalizedResourceType -eq $__prefix -or $__normalizedResourceType.StartsWith($__prefix + '/')) {
                                $__isDirectAi = $true
                                break
                            }
                        }
                        foreach ($__providerPrefix in $aiProviderPrefixes) {
                            if ($__normalizedResourceType.StartsWith($__providerPrefix)) {
                                $__isProviderAi = $true
                                break
                            }
                        }
                    }
                    if ($__isDirectAi -or $__isProviderAi) {
                        # Match the old AI subset's depth while serializing only
                        # the selected resource, never the full tenant payload.
                        $__aiJson = ConvertTo-Json -InputObject $__r -Depth 20 -Compress
                        if ($__isDirectAi) {
                            if (-not $__aiDirectFirst) { $__aiDirectWriter.Write(',') }
                            $__aiDirectWriter.Write($__aiJson)
                            $__aiDirectFirst = $false
                            $aiDirectResourceCount++
                        }
                        if ($__isProviderAi) {
                            if (-not $__aiProviderFirst) { $__aiProviderWriter.Write(',') }
                            $__aiProviderWriter.Write($__aiJson)
                            $__aiProviderFirst = $false
                            $aiProviderResourceCount++
                        }
                        $__aiJson = $null
                    }
                    # CRITICAL: drop the heavy reference. After this line, $r is
                    # unreachable (assuming no external refs) and will be collected
                    # at the next GC. This is the step that the old patch got wrong.
                    $Resources[$__ri] = $null
                    $__r = $null
                    $__json = $null
                }
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                [System.GC]::Collect()
                Write-ARIModuleMemoryCheckpoint -Stage 'after-streaming-chunk' -ModuleName ("chunk=$($__ci + 1)/$__chunkCount")
            }
        }
        $__writer.Write(']')
        $__aiDirectWriter.Write(']')
        $__aiProviderWriter.Write(']')
    } finally {
        foreach ($__streamWriter in @($__writer, $__aiDirectWriter, $__aiProviderWriter)) {
            if ($null -ne $__streamWriter) {
                try { $__streamWriter.Flush() } catch {}
                try { $__streamWriter.Dispose() } catch {}
            }
        }
    }
    if ($aiDirectResourceCount -gt 0) {
        $__aiSelectedPayloadPath = $aiDirectPayloadPath
        $__aiSelectedResourceCount = $aiDirectResourceCount
        $__aiSelectionMode = 'direct-type'
    } else {
        $__aiSelectedPayloadPath = $aiProviderPayloadPath
        $__aiSelectedResourceCount = $aiProviderResourceCount
        $__aiSelectionMode = 'provider-fallback'
    }
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'[ARI_AI_RESOURCE_SCOPE_V2] Streamed AI payload selected: mode='+$__aiSelectionMode+' count='+$__aiSelectedResourceCount)
    Write-ARIModuleMemoryCheckpoint -Stage 'after-resources-json-written'

    Remove-Variable -Name Resources
    Clear-ARIMemory

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting to Create Jobs to Process the Resources.')

    #Foreach ($ModuleFolder in $ModuleFolders)
    $ModuleFolders | ForEach-Object -Process {
            $ModuleFolder = $_
            $ModulePath = Join-Path $ModuleFolder.FullName '*.ps1'
            $ModuleName = $ModuleFolder.Name
            $ModuleFiles = Get-ChildItem -Path $ModulePath

            Write-ARIModuleMemoryCheckpoint -Stage 'before-module-job-start' -ModuleName $ModuleName
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Creating Job: '+$ModuleName)

            $c = (($JobLoop / $TotalFolders) * 100)
            $c = [math]::Round($c)
            Write-Progress -Id 1 -activity "Creating Jobs" -Status "$c% Complete." -PercentComplete $c

            $JobResourcesPayloadPath = $ResourcesJsonPath
            # ARI_AI_RESOURCE_SCOPE_V2: pass AI-relevant resources with robust type matching.
            if ($ModuleName -eq 'AI') {
                $JobResourcesPayloadPath = $__aiSelectedPayloadPath
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'[ARI_AI_RESOURCE_SCOPE_V2] AI scope selected: '+$__aiSelectionMode+' ('+$__aiSelectedResourceCount+' item(s))')
            }

            Start-Job -Name ('ResourceJob_'+$ModuleName) -ScriptBlock {

                $ModuleFiles = $($args[0])
                $Subscriptions = $($args[2])
                $InTag = $($args[3])
                $Resources = Get-Content -Path $($args[4]) -Raw -ErrorAction Stop | ConvertFrom-Json
                $Retirements = $($args[5])
                $Task = $($args[6])
                $Unsupported = $($args[10])

                $job = @()

                Foreach ($Module in $ModuleFiles)
                    {
                        $ModuleFileContent = New-Object System.IO.StreamReader($Module.FullName)
                        $ModuleData = $ModuleFileContent.ReadToEnd()
                        $ModuleFileContent.Dispose()
                        $ModName = $Module.Name.replace(".ps1","")

                        New-Variable -Name ('ModRun' + $ModName)
                        New-Variable -Name ('ModJob' + $ModName)

                        Set-Variable -Name ('ModRun' + $ModName) -Value ([PowerShell]::Create()).AddScript($ModuleData).AddArgument($PSScriptRoot).AddArgument($Subscriptions).AddArgument($InTag).AddArgument($Resources).AddArgument($Retirements).AddArgument($Task).AddArgument($null).AddArgument($null).AddArgument($null).AddArgument($Unsupported)

                        Set-Variable -Name ('ModJob' + $ModName) -Value ((get-variable -name ('ModRun' + $ModName)).Value).BeginInvoke()

                        $job += (get-variable -name ('ModJob' + $ModName)).Value
                        Start-Sleep -Milliseconds 100
                        Remove-Variable -Name ModName
                    }

                While ($Job.Runspace.IsCompleted -contains $false) { Start-Sleep -Milliseconds 500 }

                Foreach ($Module in $ModuleFiles)
                    {
                        $ModName = $Module.Name.replace(".ps1","")
                        New-Variable -Name ('ModValue' + $ModName)
                        Set-Variable -Name ('ModValue' + $ModName) -Value (((get-variable -name ('ModRun' + $ModName)).Value).EndInvoke((get-variable -name ('ModJob' + $ModName)).Value))

                        Remove-Variable -Name ('ModRun' + $ModName)
                        Remove-Variable -Name ('ModJob' + $ModName)
                        Start-Sleep -Milliseconds 100
                        Remove-Variable -Name ModName
                    }

                $Hashtable = New-Object System.Collections.Hashtable

                Foreach ($Module in $ModuleFiles)
                    {
                        $ModName = $Module.Name.replace(".ps1","")

                        $Hashtable["$ModName"] = (get-variable -name ('ModValue' + $ModName)).Value

                        Remove-Variable -Name ('ModValue' + $ModName)
                        Start-Sleep -Milliseconds 100

                        Remove-Variable -Name ModName
                    }

                $Hashtable

            } -ArgumentList $ModuleFiles, $PSScriptRoot, $Subscriptions, $InTag, $JobResourcesPayloadPath , $Retirements, 'Processing', $null, $null, $null, $Unsupported | Out-Null
            Write-ARIModuleMemoryCheckpoint -Stage 'after-module-job-start' -ModuleName $ModuleName

        if($JobLoop -eq $EnvSizeLooper)
            {
                Write-ARIModuleMemoryCheckpoint -Stage 'before-wait-resource-batch'
                Write-Host 'Waiting Batch Jobs' -ForegroundColor Cyan -NoNewline
                Write-Host '. This step may take several minutes to finish' -ForegroundColor Cyan

                $InterJobNames = Get-Job | Where-Object {
                        $_ -and
                        $_.PSObject.Properties.Match('Name').Count -gt 0 -and
                        $_.Name -like 'ResourceJob_*' -and
                        $_.PSObject.Properties.Match('State').Count -gt 0 -and
                        $_.State -eq 'Running'
                    } | ForEach-Object { $_.Name }
                # Ensure InterJobNames is always an array
                if ($null -eq $InterJobNames) {
                    $InterJobNames = @()
                } elseif ($InterJobNames -isnot [System.Array]) {
                    $InterJobNames = @($InterJobNames)
                }

                Wait-ARIJob -JobNames $InterJobNames -JobType 'Resource Batch' -LoopTime 5
                Write-ARIModuleMemoryCheckpoint -Stage 'after-wait-resource-batch'

                $JobNames = Get-Job | Where-Object {
                        $_ -and
                        $_.PSObject.Properties.Match('Name').Count -gt 0 -and
                        $_.Name -like 'ResourceJob_*'
                    } | ForEach-Object { $_.Name }
                # Ensure JobNames is always an array (PowerShell Core returns single string for one job)
                if ($null -eq $JobNames) {
                    $JobNames = @()
                } elseif ($JobNames -isnot [System.Array]) {
                    $JobNames = @($JobNames)
                }

                Build-ARICacheFiles -DefaultPath $DefaultPath -JobNames $JobNames
                Write-ARIModuleMemoryCheckpoint -Stage 'after-build-cache-files'
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                [System.GC]::Collect()
                Write-ARIModuleMemoryCheckpoint -Stage 'after-post-build-gc'

                $JobLoop = 0
            }
        $JobLoop ++

        }

        if (Test-Path -LiteralPath $ResourcesJsonPath) {
            Remove-Item -LiteralPath $ResourcesJsonPath -Force -ErrorAction SilentlyContinue
        }
        Remove-Variable -Name ResourcesJsonPath -ErrorAction SilentlyContinue
        Get-ChildItem -Path $DefaultPath -Filter 'ari-resources-ai-*.json' -File -ErrorAction SilentlyContinue | ForEach-Object {
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
        }
        Clear-ARIMemory
}
