<#
.Synopsis
Module responsible for retrieving Azure VM quotas.

.DESCRIPTION
This module retrieves Azure VM quotas for specific subscriptions and locations.

.Link
https://github.com/microsoft/ARI/Modules/Private/1.ExtractionFunctions/ResourceDetails/Get-ARIVMQuotas.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI).

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>
function Get-AriVMQuotas {
    Param ($Subscriptions, $Resources)
    $Quotas = Foreach($Sub in $Subscriptions)
        {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Getting VM Quota Details: '+$Sub.name)
            # Defensive property access: $Resources is a mix of raw ARG results
            # (with lowercase .subscriptionId / .Type) and REST API results
            # (ResourceHealth events, ManagedIdentities, AdvisorScore,
            # ReservationRecomen) which have different shapes and may not
            # carry a .subscriptionId property at all. Accessing a missing
            # property on a PSCustomObject is a terminating error in strict
            # mode and aborts the entire batch. Guard with PSObject.Properties
            # before dereferencing, matching the defensive style used elsewhere
            # in Start-ARIExtractionOrchestration.ps1.
            # Null-safe Group-Object access: when the Where-Object filter
            # eliminates all resources (e.g. subscription has zero VMs/VMSS),
            # Group-Object returns $null instead of @(). Under Set-StrictMode,
            # accessing .name on $null throws 'The property 'name' cannot be
            # found on this object' and aborts the entire batch. Capture the
            # group first, then null-check before dereferencing.
            $__VMGroups = $Resources | Where-Object {
                $hasSub  = $_.PSObject.Properties['subscriptionId']
                $hasType = $_.PSObject.Properties['Type']
                if (-not $hasSub -or -not $hasType) { return $false }
                ($_.subscriptionId -eq $Sub.id) -and ($_.Type -in 'microsoft.compute/virtualmachines','microsoft.compute/virtualmachinescalesets')
            } | Group-Object -Property Location
            if ($null -ne $__VMGroups) {
                $Locs = $__VMGroups.name
            } else {
                $Locs = $null
            }
            if (![string]::IsNullOrEmpty($Locs))
                {
                    Foreach($Loc in $Locs)
                        {
                            # Safely check Loc count - handle null/empty cases
                            $locCount = if ($null -ne $Loc -and $Loc -is [System.Array]) { $Loc.Count } elseif ($null -ne $Loc) { 1 } else { 0 }
                            if($locCount -eq 1)
                                {
                                    Set-AzContext -Subscription $Sub.Id -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -Debug:$false
                                    $Quota = get-azvmusage -location $Loc -Debug:$false
                                    $Quota = $Quota | Where-Object {$_.CurrentValue -ge 1}
                                    $tmp = [PSCustomObject]@{
                                        Location = $Loc
                                        SubId = $Sub.id
                                        Subscription = $Sub.name
                                        Data = $Quota
                                    }
                                    $tmp
                                }
                            else {
                                    Set-AzContext -Subscription $Sub.Id -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue -Debug:$false
                                    foreach($Loc1 in $Loc)
                                        {
                                            $Quota = get-azvmusage -location $Loc1 -Debug:$false
                                            $Quota = $Quota | Where-Object {$_.CurrentValue -ge 1}
                                            $tmp = [PSCustomObject]@{
                                                Location = $Loc1
                                                SubId = $Sub.id
                                                Subscription = $Sub.name
                                                Data = $Quota
                                            }
                                            $tmp
                                        }
                            }
                        }
                }
        }

    $VMQuotas = [PSCustomObject]@{
        'type'          = 'ARI/VM/Quotas'
        'properties'    = $Quotas
    }

    return $VMQuotas
}
