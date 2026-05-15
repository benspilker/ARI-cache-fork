<#
.Synopsis
Module responsible for creating the local cache files for the report.

.DESCRIPTION
This module receives the job names for the Azure Resources that were processed previously and creates the local cache files that will be used to build the Excel report.

.Link
https://github.com/microsoft/ARI/Modules/Private/2.ProcessingFunctions/Build-ARICacheFiles.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI).

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Build-ARICacheFiles {
    Param($DefaultPath, $JobNames)

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Checking Cache Folder.')

    # Ensure JobNames is always an array for safe .Count access
    if ($JobNames -isnot [System.Array]) {
        $JobNames = @($JobNames)
    }
    
    # Safely get count
    $Lops = if ($null -ne $JobNames -and $JobNames -is [System.Array]) { $JobNames.Count } elseif ($null -ne $JobNames) { 1 } else { 0 }
    $Counter = 0

    Foreach ($Job in $JobNames)
        {
            $c = (($Counter / $Lops) * 100)
            $c = [math]::Round($c)
            Write-Progress -Id 1 -activity "Building Cache Files" -Status "$c% Complete." -PercentComplete $c

            $NewJobName = ($Job -replace 'ResourceJob_','')
            $TempJob = Receive-Job -Name $Job
            $tempJobHasValues = $false
            if ($null -ne $TempJob) {
                if ($TempJob -is [System.Collections.IDictionary]) {
                    if ($TempJob.Contains("values")) {
                        $tempJobHasValues = -not [string]::IsNullOrEmpty([string]$TempJob["values"])
                    } else {
                        $tempJobHasValues = ($TempJob.Count -gt 0)
                    }
                } elseif ($TempJob.PSObject -and ($TempJob.PSObject.Properties.Name -contains "values")) {
                    $tempJobHasValues = -not [string]::IsNullOrEmpty([string]$TempJob.values)
                } elseif ($TempJob -is [System.Array]) {
                    $tempJobHasValues = ($TempJob.Count -gt 0)
                } else {
                    $tempJobHasValues = $true
                }
            }
            if ($tempJobHasValues)
                {
                    $JobJSONName = ($NewJobName+'.json')
                    $JobFileName = Join-Path $DefaultPath 'ReportCache' $JobJSONName
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Creating Cache File: '+ $JobFileName)

                    $jobObjectToWrite = $TempJob
                    if ($NewJobName -eq 'AI') {
                        # ARI_AI_CACHE_MEMORY_SAFE: trim remoting metadata and serialize AI cache with a lower depth to reduce peak memory.
                        if ($jobObjectToWrite -and $jobObjectToWrite.PSObject) {
                            foreach ($metaProp in @('PSComputerName','RunspaceId','PSShowComputerName')) {
                                try { [void]$jobObjectToWrite.PSObject.Properties.Remove($metaProp) } catch {}
                            }
                        }
                    }
                    $jsonDepth = if ($NewJobName -eq 'AI') { 12 } else { 40 }
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    [System.GC]::Collect()
                    $jobObjectToWrite | ConvertTo-Json -Depth $jsonDepth -Compress | Set-Content -Path $JobFileName -Encoding UTF8
                    Remove-Variable -Name jobObjectToWrite -ErrorAction SilentlyContinue
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    [System.GC]::Collect()
                }
            elseif ($NewJobName -eq 'AI')
                {
                    # Always emit deterministic AI.json even when there are no AI rows.
                    $JobJSONName = ($NewJobName+'.json')
                    $JobFileName = Join-Path $DefaultPath 'ReportCache' $JobJSONName
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Creating empty AI cache file: '+ $JobFileName)
                    '[]' | Set-Content -Path $JobFileName -Encoding UTF8
                }
            Remove-Job -Name $Job
            Remove-Variable -Name TempJob -ErrorAction SilentlyContinue

            $Counter++

        }
    Clear-ARIMemory
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Cache Files Created.')
}
