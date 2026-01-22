<#
.Synopsis
Extraction orchestration for Azure Resource Inventory

.DESCRIPTION
This module orchestrates the extraction of resources for Azure Resource Inventory.

.Link
https://github.com/microsoft/ARI/Modules/Private/0.MainFunctions/Start-ARIExtractionOrchestration.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.11
First Release Date: 15th Oct, 2024
Authors: Claudio Merola

#>
function Start-ARIExtractionOrchestration {
    Param($ManagementGroup, $Subscriptions, $SubscriptionID, $SkipPolicy, $ResourceGroup, $SecurityCenter, $SkipAdvisory, $IncludeTags, $TagKey, $TagValue, $SkipAPIs, $SkipVMDetails, $IncludeCosts, $Automation, $AzureEnvironment)

    $GraphData = Start-ARIGraphExtraction -ManagementGroup $ManagementGroup -Subscriptions $Subscriptions -SubscriptionID $SubscriptionID -ResourceGroup $ResourceGroup -SecurityCenter $SecurityCenter -SkipAdvisory $SkipAdvisory -IncludeTags $IncludeTags -TagKey $TagKey -TagValue $TagValue -AzureEnvironment $AzureEnvironment

    # Initialize all variables as arrays to prevent "variable not set" errors
    $Resources = if ($null -ne $GraphData.Resources) { $GraphData.Resources } else { @() }
    $ResourceContainers = if ($null -ne $GraphData.ResourceContainers) { $GraphData.ResourceContainers } else { @() }
    $Advisories = if ($null -ne $GraphData.Advisories) { $GraphData.Advisories } else { @() }
    $Security = if ($null -ne $GraphData.Security) { $GraphData.Security } else { @() }
    $Retirements = if ($null -ne $GraphData.Retirements) { $GraphData.Retirements } else { @() }
    
    # Initialize optional variables that may not be set in all code paths
    $VMQuotas = $null
    $VMSkuDetails = $null
    $Costs = $null
    $PolicyAssign = $null
    $PolicyDef = $null
    $PolicySetDef = $null
    
    # Ensure Resources is always an array (not a single value) for += operations
    if ($Resources -isnot [System.Array]) {
        $Resources = @($Resources)
    }

    Remove-Variable -Name GraphData -ErrorAction SilentlyContinue

    # Safely access Count properties - handle null/empty cases
    # Check if variables are arrays before accessing .Count
    $ResourcesCount = if ($null -ne $Resources -and $Resources -is [System.Array]) { [string]$Resources.Count } elseif ($null -ne $Resources) { "1" } else { "0" }
    $AdvisoryCount = if ($null -ne $Advisories -and $Advisories -is [System.Array]) { [string]$Advisories.Count } elseif ($null -ne $Advisories) { "1" } else { "0" }
    $SecCenterCount = if ($null -ne $Security -and $Security -is [System.Array]) { [string]$Security.Count } elseif ($null -ne $Security) { "1" } else { "0" }

    if(!$SkipAPIs.IsPresent)
        {
            Write-Progress -activity 'Azure Inventory' -Status "12% Complete." -PercentComplete 12 -CurrentOperation "Starting API Extraction.."
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Getting API Resources.')
            $APIResults = Get-ARIAPIResources -Subscriptions $Subscriptions -AzureEnvironment $AzureEnvironment -SkipPolicy $SkipPolicy
            $Resources += $APIResults.ResourceHealth
            $Resources += $APIResults.ManagedIdentities
            $Resources += $APIResults.AdvisorScore
            $Resources += $APIResults.ReservationRecomen
            $PolicyAssign = $APIResults.PolicyAssign
            $PolicyDef = $APIResults.PolicyDef
            $PolicySetDef = $APIResults.PolicySetDef
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'API Resource Inventory Finished.')
            Remove-Variable APIResults -ErrorAction SilentlyContinue
        }

    # Safely access PolicyAssign.policyAssignments.Count - handle null/empty cases
    if ($null -ne $PolicyAssign -and $null -ne $PolicyAssign.policyAssignments) {
        if ($PolicyAssign.policyAssignments -is [System.Array]) {
            $PolicyCount = [string]$PolicyAssign.policyAssignments.Count
        } else {
            $PolicyCount = "1"
        }
    } else {
        $PolicyCount = "0"
    }

    if ($IncludeCosts.IsPresent) {
        $Costs = Get-ARICostInventory -Subscriptions $Subscriptions -Days 60 -Granularity 'Monthly'
    }

    if (!$SkipVMDetails.IsPresent)
        {
            Write-Host 'Gathering VM Extra Details: ' -NoNewline
            Write-Host 'Quotas' -ForegroundColor Cyan
            Write-Progress -activity 'Azure Inventory' -Status "13% Complete." -PercentComplete 13 -CurrentOperation "Starting VM Details Extraction.."

            $VMQuotas = Get-AriVMQuotas -Subscriptions $Subscriptions -Resources $Resources

            $Resources += $VMQuotas

            # Don't remove VMQuotas - it's needed in the return object

            Write-Host 'Gathering VM Extra Details: ' -NoNewline
            Write-Host 'Size SKU' -ForegroundColor Cyan

            $VMSkuDetails = Get-ARIVMSkuDetails -Resources $Resources

            $Resources += $VMSkuDetails

            Remove-Variable -Name VMSkuDetails -ErrorAction SilentlyContinue

        }

    $ReturnData = [PSCustomObject]@{
        Resources = $Resources
        Quotas = $VMQuotas
        Costs = $Costs
        ResourceContainers = $ResourceContainers
        Advisories = $Advisories
        ResourcesCount = $ResourcesCount
        AdvisoryCount = $AdvisoryCount
        SecCenterCount = $SecCenterCount
        Security = $Security
        Retirements = $Retirements
        PolicyCount = $PolicyCount
        PolicyAssign = $PolicyAssign
        PolicyDef = $PolicyDef
        PolicySetDef = $PolicySetDef
    }

    return $ReturnData
}