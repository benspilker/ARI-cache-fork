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

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Converting Resource data to JSON for Jobs')
    $ResourcesJsonPath = Join-Path $DefaultPath ("ari-resources-" + [guid]::NewGuid().ToString("N") + ".json")
    $Resources | ConvertTo-Json -Depth 40 -Compress | Set-Content -Path $ResourcesJsonPath -Encoding UTF8
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
                $aiResourcesPayloadPath = Join-Path $DefaultPath ("ari-resources-ai-" + [guid]::NewGuid().ToString("N") + ".json")
                try {
                    $aiResourcesRaw = Get-Content -Path $ResourcesJsonPath -Raw -ErrorAction Stop | ConvertFrom-Json
                    $allResources = @($aiResourcesRaw)
                    $aiResources = @($allResources | Where-Object {
                        $rtype = ''
                        try {
                            if ($_.PSObject.Properties['type']) { $rtype = [string]$_.type }
                            elseif ($_.PSObject.Properties['TYPE']) { $rtype = [string]$_.TYPE }
                            elseif ($_.PSObject.Properties['resourceType']) { $rtype = [string]$_.resourceType }
                            elseif ($_.PSObject.Properties['ResourceType']) { $rtype = [string]$_.ResourceType }
                        } catch {}
                        if ([string]::IsNullOrWhiteSpace($rtype)) { return $false }
                        $rtypeNormalized = $rtype.Trim().ToLowerInvariant()
                        foreach ($prefix in $aiResourceTypePrefixes) {
                            if ($rtypeNormalized -eq $prefix -or $rtypeNormalized.StartsWith($prefix + '/')) {
                                return $true
                            }
                        }
                        return $false
                    })

                    if ($aiResources.Count -eq 0 -and $allResources.Count -gt 0) {
                        $providerMatches = @($allResources | Where-Object {
                            $rtype = ''
                            try {
                                if ($_.PSObject.Properties['type']) { $rtype = [string]$_.type }
                                elseif ($_.PSObject.Properties['TYPE']) { $rtype = [string]$_.TYPE }
                                elseif ($_.PSObject.Properties['resourceType']) { $rtype = [string]$_.resourceType }
                                elseif ($_.PSObject.Properties['ResourceType']) { $rtype = [string]$_.ResourceType }
                            } catch {}
                            if ([string]::IsNullOrWhiteSpace($rtype)) { return $false }
                            $rtypeNormalized = $rtype.Trim().ToLowerInvariant()
                            foreach ($providerPrefix in $aiProviderPrefixes) {
                                if ($rtypeNormalized.StartsWith($providerPrefix)) { return $true }
                            }
                            return $false
                        })
                        if ($providerMatches.Count -gt 0) {
                            $aiResources = $providerMatches
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'[PATCHED] AI scope fallback used provider-prefix matches: '+$($aiResources.Count)+' item(s)')
                        }
                    }

                    $aiResources | ConvertTo-Json -Depth 20 -Compress | Set-Content -Path $aiResourcesPayloadPath -Encoding UTF8
                    $JobResourcesPayloadPath = $aiResourcesPayloadPath
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'[PATCHED] AI resource scope applied: '+$($aiResources.Count)+' item(s) from total '+$($allResources.Count))
                } catch {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'[PATCHED] AI resource scope fallback to full payload due to error: '+$_.Exception.Message)
                    $JobResourcesPayloadPath = $ResourcesJsonPath
                }
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
