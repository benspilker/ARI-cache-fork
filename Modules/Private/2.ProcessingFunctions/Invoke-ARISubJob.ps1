<#
.Synopsis
Module responsible for invoking subscription processing jobs.

.DESCRIPTION
This module starts jobs to process Azure subscriptions and their associated resources, either in automation or manual mode.

.Link
https://github.com/microsoft/ARI/Modules/Private/2.ProcessingFunctions/Invoke-ARISubJob.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI).

.NOTES
Version: 3.6.5
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Invoke-ARISubJob {
    Param($Subscriptions, $Automation, $Resources, $CostData, $ARIModule)
    
    # Calculate function file path based on ARIModule location
    # If ARIModule is a path, use it; if it's a module name, find the module
    $functionFilePath = $null
    
    # Check if ARIModule looks like a path (contains path separators or file extension)
    $isPath = $ARIModule -match '[\\/]|\.psm1$|\.psd1$'
    
    if ($isPath) {
        # ARIModule appears to be a path - resolve it
        try {
            $resolvedPath = Resolve-Path $ARIModule -ErrorAction SilentlyContinue
            if ($null -ne $resolvedPath) {
                $moduleDir = Split-Path $resolvedPath.Path -Parent
                $functionFilePath = Join-Path $moduleDir "Modules\Public\PublicFunctions\Jobs\Start-ARISubscriptionJob.ps1"
            }
        } catch {
            # Path resolution failed, try as-is
            if (Test-Path $ARIModule -ErrorAction SilentlyContinue) {
                $moduleDir = Split-Path $ARIModule -Parent
                $functionFilePath = Join-Path $moduleDir "Modules\Public\PublicFunctions\Jobs\Start-ARISubscriptionJob.ps1"
            }
        }
    } else {
        # ARIModule is a module name - try to find the module
        try {
            # First try loaded modules
            $moduleInfo = Get-Module -Name $ARIModule -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($null -eq $moduleInfo) {
                # Try available modules
                $moduleInfo = Get-Module -Name $ARIModule -ListAvailable -ErrorAction SilentlyContinue | Select-Object -First 1
            }
            if ($null -ne $moduleInfo -and $null -ne $moduleInfo.ModuleBase) {
                $functionFilePath = Join-Path $moduleInfo.ModuleBase "Modules\Public\PublicFunctions\Jobs\Start-ARISubscriptionJob.ps1"
            }
        } catch {
            # Module lookup failed, will try fallback
        }
    }
    
    # If we still don't have a path, try to find it relative to the current script location
    if ($null -eq $functionFilePath -or -not (Test-Path $functionFilePath -ErrorAction SilentlyContinue)) {
        # Try to find it relative to the current script location
        try {
            $currentScriptDir = (Get-Item $PSScriptRoot).Parent.Parent.Parent
            $functionFilePath = Join-Path $currentScriptDir "Modules\Public\PublicFunctions\Jobs\Start-ARISubscriptionJob.ps1"
        } catch {
            # If that fails, set to null - the job will try to use module import only
            $functionFilePath = $null
        }
    }
    
    # Log the resolved path for debugging (but don't fail if it's null)
    if ($null -ne $functionFilePath) {
        Write-Debug "Invoke-ARISubJob: Resolved function path: $functionFilePath (exists: $(Test-Path $functionFilePath -ErrorAction SilentlyContinue))"
    } else {
        Write-Debug "Invoke-ARISubJob: Could not resolve function path - will rely on module import only"
    }

    if ($Automation.IsPresent)
        {
            Write-Output ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Subscription Job')
            Start-ThreadJob -Name 'Subscriptions' -ScriptBlock {

                param($SubsParam, $ResParam, $ModulePath, $CostDataParam, $FunctionPath)
                
                # Import module first
                Import-Module $ModulePath -Force -DisableNameChecking -ErrorAction Stop
                
                # Explicitly dot-source the function file to ensure it's available
                # This is a fallback in case module import doesn't load the function
                if (Test-Path $FunctionPath) {
                    . $FunctionPath
                }
                
                # Verify function exists
                if (-not (Get-Command Start-ARISubscriptionJob -ErrorAction SilentlyContinue)) {
                    throw "Start-ARISubscriptionJob function not found after module import and dot-sourcing"
                }

                # Ensure variables are arrays after job serialization (arrays can become null or lose type)
                if ($null -eq $SubsParam) {
                    $SubsParam = @()
                } elseif ($SubsParam -isnot [System.Array]) {
                    $SubsParam = @($SubsParam)
                }
                
                if ($null -eq $ResParam) {
                    $ResParam = @()
                } elseif ($ResParam -isnot [System.Array]) {
                    $ResParam = @($ResParam)
                }

                $SubResult = Start-ARISubscriptionJob -Subscriptions $SubsParam -Resources $ResParam -CostData $CostDataParam
                
                # Debug: Log what was returned (use Write-Information to avoid interfering with output stream)
                $resultCount = if ($null -ne $SubResult -and $SubResult -is [System.Array]) { $SubResult.Count } elseif ($null -ne $SubResult) { 1 } else { 0 }
                Write-Information "Invoke-ARISubJob: Start-ARISubscriptionJob returned $resultCount result(s)" -InformationAction SilentlyContinue

                # Explicitly output the result using Write-Output to ensure it's captured by Receive-Job
                Write-Output $SubResult

            } -ArgumentList $Subscriptions, $Resources, $ARIModule, $CostData, $functionFilePath | Out-Null
        }
    else
        {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Subscription Job.')
            
            # Resolve module path before passing to job
            $resolvedModulePath = $ARIModule
            if ($ARIModule -match '[\\/]|\.psm1$|\.psd1$') {
                # It's a path - resolve it
                try {
                    $resolvedPath = Resolve-Path $ARIModule -ErrorAction Stop
                    $resolvedModulePath = $resolvedPath.Path
                } catch {
                    # If resolution fails, try as-is
                    if (Test-Path $ARIModule -ErrorAction SilentlyContinue) {
                        $resolvedModulePath = (Resolve-Path $ARIModule).Path
                    }
                }
            }
            
            Start-Job -Name 'Subscriptions' -ScriptBlock {

                param($SubsParam, $ResParam, $ModulePath, $CostDataParam, $FunctionPath)
                
                # Import module first
                Import-Module $ModulePath -Force -DisableNameChecking -ErrorAction Stop
                
                # Explicitly dot-source the function file to ensure it's available
                # This is a fallback in case module import doesn't load the function
                if ($null -ne $FunctionPath -and (Test-Path $FunctionPath -ErrorAction SilentlyContinue)) {
                    . $FunctionPath
                } else {
                    # If function path wasn't provided or doesn't exist, try to find it
                    # Get the module that was just imported
                    $moduleInfo = Get-Module -Name 'AzureResourceInventory' -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($null -ne $moduleInfo -and $null -ne $moduleInfo.ModuleBase) {
                        $functionFilePath = Join-Path $moduleInfo.ModuleBase "Modules\Public\PublicFunctions\Jobs\Start-ARISubscriptionJob.ps1"
                        if (Test-Path $functionFilePath -ErrorAction SilentlyContinue) {
                            . $functionFilePath
                        }
                    }
                }
                
                # Verify function exists - if not, try one more time to dot-source from module base
                if (-not (Get-Command Start-ARISubscriptionJob -ErrorAction SilentlyContinue)) {
                    # Last resort: try to find and dot-source the function file
                    $moduleInfo = Get-Module -Name 'AzureResourceInventory' -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($null -ne $moduleInfo -and $null -ne $moduleInfo.ModuleBase) {
                        $altFunctionPath = Join-Path $moduleInfo.ModuleBase "Modules\Public\PublicFunctions\Jobs\Start-ARISubscriptionJob.ps1"
                        if (Test-Path $altFunctionPath -ErrorAction SilentlyContinue) {
                            . $altFunctionPath
                        }
                    }
                    
                    # If still not found, that's OK - module import should have loaded it
                    # Don't throw - let the function call fail naturally if it doesn't exist
                }

                # Ensure variables are arrays after job serialization (arrays can become null or lose type)
                if ($null -eq $SubsParam) {
                    $SubsParam = @()
                } elseif ($SubsParam -isnot [System.Array]) {
                    $SubsParam = @($SubsParam)
                }
                
                if ($null -eq $ResParam) {
                    $ResParam = @()
                } elseif ($ResParam -isnot [System.Array]) {
                    $ResParam = @($ResParam)
                }

                $SubResult = Start-ARISubscriptionJob -Subscriptions $SubsParam -Resources $ResParam -CostData $CostDataParam
                
                # Debug: Log what was returned (use Write-Information to avoid interfering with output stream)
                $resultCount = if ($null -ne $SubResult -and $SubResult -is [System.Array]) { $SubResult.Count } elseif ($null -ne $SubResult) { 1 } else { 0 }
                Write-Information "Invoke-ARISubJob: Start-ARISubscriptionJob returned $resultCount result(s)" -InformationAction SilentlyContinue

                # Explicitly output the result using Write-Output to ensure it's captured by Receive-Job
                Write-Output $SubResult

            } -ArgumentList $Subscriptions, $Resources, $ARIModule, $CostData, $functionFilePath | Out-Null
        }

}
