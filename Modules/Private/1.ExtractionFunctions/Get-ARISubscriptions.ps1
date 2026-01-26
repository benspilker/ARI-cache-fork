<#
.Synopsis
Module responsible for retrieving Azure subscriptions.

.DESCRIPTION
This module retrieves Azure subscriptions for a given tenant or specific subscription IDs.

.Link
https://github.com/microsoft/ARI/Modules/Private/1.ExtractionFunctions/Get-ARISubscriptions.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI).

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>
function Get-ARISubscriptions {
    Param ($TenantID,$SubscriptionID,$PlatOS)
    function Invoke-QuietAzIdentityLogging {
        param([scriptblock]$ScriptBlock)

        $prevEnabled = $env:AZURE_IDENTITY_LOGGING_ENABLED
        $prevLevel = $env:AZURE_IDENTITY_LOGGING_LEVEL
        $prevDebug = $DebugPreference
        $prevVerbose = $VerbosePreference
        $prevInfo = $InformationPreference
        $prevAzDebug = $null
        $prevAzAccountsDebug = $null
        $hadAzDebug = $PSDefaultParameterValues.ContainsKey('Az.*:Debug')
        $hadAzAccountsDebug = $PSDefaultParameterValues.ContainsKey('Az.Accounts.*:Debug')
        if ($hadAzDebug) { $prevAzDebug = $PSDefaultParameterValues['Az.*:Debug'] }
        if ($hadAzAccountsDebug) { $prevAzAccountsDebug = $PSDefaultParameterValues['Az.Accounts.*:Debug'] }
        try {
            $env:AZURE_IDENTITY_LOGGING_ENABLED = 'false'
            $env:AZURE_IDENTITY_LOGGING_LEVEL = 'warning'
            $DebugPreference = 'SilentlyContinue'
            $VerbosePreference = 'SilentlyContinue'
            $InformationPreference = 'SilentlyContinue'
            $PSDefaultParameterValues['Az.*:Debug'] = $false
            $PSDefaultParameterValues['Az.Accounts.*:Debug'] = $false
            & $ScriptBlock
        } finally {
            $DebugPreference = $prevDebug
            $VerbosePreference = $prevVerbose
            $InformationPreference = $prevInfo
            if ($hadAzDebug) { $PSDefaultParameterValues['Az.*:Debug'] = $prevAzDebug } else { $PSDefaultParameterValues.Remove('Az.*:Debug') | Out-Null }
            if ($hadAzAccountsDebug) { $PSDefaultParameterValues['Az.Accounts.*:Debug'] = $prevAzAccountsDebug } else { $PSDefaultParameterValues.Remove('Az.Accounts.*:Debug') | Out-Null }
            if ($null -ne $prevEnabled) { $env:AZURE_IDENTITY_LOGGING_ENABLED = $prevEnabled } else { Remove-Item Env:AZURE_IDENTITY_LOGGING_ENABLED -ErrorAction SilentlyContinue }
            if ($null -ne $prevLevel) { $env:AZURE_IDENTITY_LOGGING_LEVEL = $prevLevel } else { Remove-Item Env:AZURE_IDENTITY_LOGGING_LEVEL -ErrorAction SilentlyContinue }
        }
    }
    if($PlatOS -eq 'Azure CloudShell')
        {
            $Subscriptions = Invoke-QuietAzIdentityLogging { Get-AzSubscription -WarningAction SilentlyContinue -Debug:$false }
            
            if ($SubscriptionID)
                {
                    # Safely check SubscriptionID count - handle null/empty cases
                    $subIdCount = if ($null -ne $SubscriptionID -and $SubscriptionID -is [System.Array]) { $SubscriptionID.Count } elseif ($null -ne $SubscriptionID) { 1 } else { 0 }
                    if($subIdCount -gt 1)
                        {
                            $Subscriptions = $Subscriptions | Where-Object { $_.ID -in $SubscriptionID }
                        }
                    elseif($subIdCount -eq 1)
                        {
                            # Handle both single string and single-element array
                            $singleSubId = if ($SubscriptionID -is [System.Array]) { $SubscriptionID[0] } else { $SubscriptionID }
                            $Subscriptions = $Subscriptions | Where-Object { $_.ID -eq $singleSubId }
                        }
                }
        }
    else
        {
            Write-Host "Extracting Subscriptions from Tenant $TenantID"
            try
                {
                    $Subscriptions = Invoke-QuietAzIdentityLogging { Get-AzSubscription -TenantId $TenantID -WarningAction SilentlyContinue -Debug:$false }
                }
            catch
                {
                    Write-Host ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+ " Error: $_")
                    return
                }
            
            if ($SubscriptionID)
                {
                    # Safely check SubscriptionID count - handle null/empty cases
                    $subIdCount = if ($null -ne $SubscriptionID -and $SubscriptionID -is [System.Array]) { $SubscriptionID.Count } elseif ($null -ne $SubscriptionID) { 1 } else { 0 }
                    if($subIdCount -gt 1)
                        {
                            $Subscriptions = $Subscriptions | Where-Object { $_.ID -in $SubscriptionID }
                        }
                    elseif($subIdCount -eq 1)
                        {
                            # Handle both single string and single-element array
                            $singleSubId = if ($SubscriptionID -is [System.Array]) { $SubscriptionID[0] } else { $SubscriptionID }
                            $Subscriptions = $Subscriptions | Where-Object { $_.ID -eq $singleSubId }
                        }
                }
        }
    
    return $Subscriptions
}
