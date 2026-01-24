<#
.Synopsis
Module for Excel Job Processing

.DESCRIPTION
This script processes inventory modules and builds the Excel report.

.Link
https://github.com/microsoft/ARI/Modules/Private/3.ReportingFunctions/Start-ARIExcelJob.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Start-ARIExcelJob {
    Param($ReportCache, $File, $TableStyle)

    Write-Host "[DEBUG] Start-ARIExcelJob: Starting with ReportCache=$ReportCache, File=$File" -ForegroundColor Magenta
    
    try {
        $ParentPath = (get-item $PSScriptRoot).parent.parent
        $InventoryModulesPath = Join-Path $ParentPath 'Public' 'InventoryModules'
        Write-Host "[DEBUG] Start-ARIExcelJob: InventoryModulesPath=$InventoryModulesPath" -ForegroundColor Magenta
        
        # Safely get module folders - ensure it's always an array
        $ModuleFolders = @(Get-ChildItem -Path $InventoryModulesPath -Directory -ErrorAction SilentlyContinue)
        if ($null -eq $ModuleFolders -or $ModuleFolders.Count -eq 0) {
            Write-Warning "No module folders found in $InventoryModulesPath"
            $ModuleFolders = @()
        }
        Write-Host "[DEBUG] Start-ARIExcelJob: Found $($ModuleFolders.Count) module folder(s)" -ForegroundColor Magenta

    Write-Progress -activity 'Azure Inventory' -Status "68% Complete." -PercentComplete 68 -CurrentOperation "Starting the Report Loop.."

    # Safely get module count - handle null or empty results
    $moduleFiles = @(Get-ChildItem -Path $InventoryModulesPath -Recurse -Filter "*.ps1" -ErrorAction SilentlyContinue)
    $ModulesCount = if ($null -ne $moduleFiles -and $moduleFiles.Count -gt 0) { [string]$moduleFiles.Count } else { "0" }

    Write-Output 'Starting to Build Excel Report.'
    Write-Host 'Supported Resource Types: ' -NoNewline -ForegroundColor Green
    Write-Host $ModulesCount -ForegroundColor Cyan

    $Lops = $ModulesCount
    $ReportCounter = 0

        Foreach ($ModuleFolder in $ModuleFolders)
        {
            $CacheData = $null
            $ModulePath = Join-Path $ModuleFolder.FullName '*.ps1'
            
            # Safely get module files - ensure it's always an array
            $ModuleFiles = @(Get-ChildItem -Path $ModulePath -ErrorAction SilentlyContinue)
            if ($null -eq $ModuleFiles) {
                $ModuleFiles = @()
            }

            # Safely get cache files - ensure it's always an array
            $CacheFiles = @(Get-ChildItem -Path $ReportCache -Recurse -ErrorAction SilentlyContinue)
            if ($null -eq $CacheFiles) {
                $CacheFiles = @()
            }
            
            $JSONFileName = ($ModuleFolder.Name + '.json')
            $CacheFile = $CacheFiles | Where-Object { $_.Name -like "*$JSONFileName" }
            
            # SPECIAL CASE: Outages module can use standalone Outages.json file
            # This allows clean separation from APIs.json and avoids merge conflicts
            if ($ModuleFolder.Name -eq 'APIs') {
                $outagesStandaloneFile = $CacheFiles | Where-Object { $_.Name -eq 'Outages.json' }
                if ($outagesStandaloneFile) {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+"[Outages] Found standalone Outages.json file")
                }
            }

            if ($CacheFile)
                {
                    $CacheFileContent = New-Object System.IO.StreamReader($CacheFile.FullName)
                    $CacheData = $CacheFileContent.ReadToEnd()
                    $CacheFileContent.Dispose()
                    $CacheData = $CacheData | ConvertFrom-Json
                }

            Foreach ($Module in $ModuleFiles)
                {
                    $c = (($ReportCounter / $Lops) * 100)
                    $c = [math]::Round($c)
                    Write-Progress -Id 1 -activity "Building Report" -Status "$c% Complete." -PercentComplete $c

                    $ModuleFileContent = New-Object System.IO.StreamReader($Module.FullName)
                    $ModuleData = $ModuleFileContent.ReadToEnd()
                    $ModuleFileContent.Dispose()
                    $ModName = $Module.Name.replace(".ps1","")

                    # Safely access cache data - check if CacheData exists and has the ModName property
                    $SmaResources = $null
                    
                    # SPECIAL CASE: Outages module can use standalone Outages.json file
                    if ($ModName -eq 'Outages') {
                        $outagesStandaloneFile = $CacheFiles | Where-Object { $_.Name -eq 'Outages.json' }
                        if ($outagesStandaloneFile) {
                            try {
                                $outagesFileContent = New-Object System.IO.StreamReader($outagesStandaloneFile.FullName)
                                $outagesFileData = $outagesFileContent.ReadToEnd()
                                $outagesFileContent.Dispose()
                                $outagesFileData = $outagesFileData | ConvertFrom-Json
                                
                                # Handle both direct array and object with Outages property
                                if ($outagesFileData -is [System.Array]) {
                                    $SmaResources = $outagesFileData
                                } elseif ($outagesFileData.PSObject.Properties.Name -contains 'Outages') {
                                    $SmaResources = $outagesFileData.Outages
                                } else {
                                    $SmaResources = @($outagesFileData)
                                }
                                
                                $outagesCount = 0
                                if ($null -ne $SmaResources) {
                                    if ($SmaResources -is [System.Array]) {
                                        $outagesCount = $SmaResources.Count
                                    } elseif ($null -ne $SmaResources) {
                                        $outagesCount = 1
                                    }
                                }
                                Write-Host "[Outages] Found $outagesCount outage(s) in standalone Outages.json file" -ForegroundColor Green
                            } catch {
                                Write-Host "[Outages] WARNING: Failed to read standalone Outages.json: $_" -ForegroundColor Yellow
                            }
                        }
                    }
                    
                    # Fallback to standard cache file lookup if Outages.json not found or for other modules
                    if ($null -eq $SmaResources -and $null -ne $CacheData) {
                        if ($CacheData.PSObject.Properties.Name -contains $ModName) {
                            $SmaResources = $CacheData.$ModName
                            # Debug logging for Outages module
                            if ($ModName -eq 'Outages') {
                                $outagesCount = 0
                                if ($null -ne $SmaResources) {
                                    if ($SmaResources -is [System.Array]) {
                                        $outagesCount = $SmaResources.Count
                                    } elseif ($null -ne $SmaResources) {
                                        $outagesCount = 1
                                    }
                                }
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+"[Outages] Found $outagesCount outage(s) in cache file $JSONFileName")
                                if ($outagesCount -eq 0) {
                                    Write-Host "[Outages] WARNING: Cache file $JSONFileName has Outages key but contains no data" -ForegroundColor Yellow
                                    Write-Host "[Outages] CacheData properties: $($CacheData.PSObject.Properties.Name -join ', ')" -ForegroundColor Gray
                                }
                            }
                        } else {
                            # Debug logging for missing module data
                            if ($ModName -eq 'Outages') {
                                Write-Host "[Outages] WARNING: Cache file $JSONFileName does not contain '$ModName' property" -ForegroundColor Yellow
                                if ($null -ne $CacheData) {
                                    Write-Host "[Outages] Available properties: $($CacheData.PSObject.Properties.Name -join ', ')" -ForegroundColor Gray
                                } else {
                                    Write-Host "[Outages] CacheData is null" -ForegroundColor Gray
                                }
                            }
                        }
                    } else {
                        if ($ModName -eq 'Outages' -and $null -eq $SmaResources) {
                            Write-Host "[Outages] WARNING: Cache file $JSONFileName not found or could not be parsed, and standalone Outages.json not found" -ForegroundColor Yellow
                        }
                    }

                    # Safely get count - handle null, array, or single object
                    $ModuleResourceCount = 0
                    if ($null -ne $SmaResources) {
                        if ($SmaResources -is [System.Array]) {
                            $ModuleResourceCount = $SmaResources.Count
                        } elseif ($SmaResources -is [System.Collections.Hashtable]) {
                            $ModuleResourceCount = $SmaResources.Count
                        } elseif ($SmaResources -is [PSCustomObject]) {
                            $ModuleResourceCount = 1
                        } else {
                            # Try to get count property
                            try {
                                $ModuleResourceCount = $SmaResources.count
                            } catch {
                                $ModuleResourceCount = 0
                            }
                        }
                    }

                    if ($ModuleResourceCount -gt 0)
                    {
                        Start-Sleep -Milliseconds 25
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+"Running Module: '$ModName'. Excel Rows: $ModuleResourceCount")

                        # Ensure the output directory exists before writing
                        $fileDir = Split-Path -Path $File -Parent
                        if (-not (Test-Path $fileDir)) {
                            New-Item -ItemType Directory -Path $fileDir -Force | Out-Null
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+"Created directory: $fileDir")
                        }

                        $ScriptBlock = [Scriptblock]::Create($ModuleData)

                        try {
                            Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $PSScriptRoot, $null, $InTag, $null, $null, 'Reporting', $File, $SmaResources, $TableStyle, $null -ErrorAction Stop
                            
                            # CRITICAL: Longer delay after Export-Excel to ensure file is fully written and closed
                            # Export-Excel with -Path opens, writes, and closes the file, but EPPlus may need time
                            # to fully release file handles before the next module can access it
                            Start-Sleep -Milliseconds 500
                            
                            # Force multiple garbage collection cycles to release any file handles
                            # This is critical for preventing "Error saving file" when modules run sequentially
                            for ($i = 1; $i -le 3; $i++) {
                                [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $false)
                                [System.GC]::WaitForPendingFinalizers()
                            }
                            [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                            
                            # Additional small delay after GC to ensure file system has released the lock
                            Start-Sleep -Milliseconds 100
                        } catch {
                            $errorMsg = $_.Exception.Message
                            Write-Error "Module '$ModName' failed: $errorMsg"
                            Write-Error "Stack trace: $($_.ScriptStackTrace)"
                            throw
                        }

                    }

                    $ReportCounter ++

                }
                Remove-Variable -Name CacheData
                Remove-Variable -Name SmaResources
                Clear-ARIMemory
        }
        Write-Progress -Id 1 -activity "Building Report" -Status "100% Complete." -Completed
        Write-Host "[DEBUG] Start-ARIExcelJob: Completed successfully" -ForegroundColor Magenta
    } catch {
        $errorMsg = "Error in Start-ARIExcelJob: $($_.Exception.Message)"
        $errorLine = if ($null -ne $_.InvocationInfo) { $_.InvocationInfo.ScriptLineNumber } else { 'Unknown' }
        $errorFunc = if ($null -ne $_.InvocationInfo -and $null -ne $_.InvocationInfo.PSObject.Properties['FunctionName']) { $_.InvocationInfo.FunctionName } else { 'Unknown' }
        $errorStack = if ($null -ne $_.ScriptStackTrace) { $_.ScriptStackTrace } else { 'No stack trace available' }
        Write-Host "[ERROR] $errorMsg" -ForegroundColor Red
        Write-Host "[ERROR] Line: $errorLine, Function: $errorFunc" -ForegroundColor Red
        Write-Host "[ERROR] Stack: $errorStack" -ForegroundColor Red
        Write-Error $errorMsg
        Write-Error "Stack trace: $errorStack"
        throw
    }
}