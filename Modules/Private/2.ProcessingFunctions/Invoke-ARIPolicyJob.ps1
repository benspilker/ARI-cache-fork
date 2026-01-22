<#
.Synopsis
Module responsible for invoking policy evaluation jobs.

.DESCRIPTION
This module starts jobs to evaluate Azure policies, including policy definitions, assignments, and set definitions, either in automation or manual mode.

.Link
https://github.com/microsoft/ARI/Modules/Private/2.ProcessingFunctions/Invoke-ARIPolicyJob.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI).

.NOTES
Version: 3.6.5
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Invoke-ARIPolicyJob {
    Param($Subscriptions, $PolicySetDef, $PolicyAssign, $PolicyDef, $ARIModule, $Automation)

    if ($Automation.IsPresent)
        {
            Write-Output ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Policy Job')
            Start-ThreadJob -Name 'Policy' -ScriptBlock {

                import-module $($args[4])

                # Ensure variables are arrays after job serialization (arrays can become null or lose type)
                $SubsParam = $($args[0])
                if ($null -eq $SubsParam) {
                    $SubsParam = @()
                } elseif ($SubsParam -isnot [System.Array]) {
                    $SubsParam = @($SubsParam)
                }
                
                $PolicySetDefParam = $($args[1])
                if ($null -eq $PolicySetDefParam) {
                    $PolicySetDefParam = @()
                } elseif ($PolicySetDefParam -isnot [System.Array]) {
                    $PolicySetDefParam = @($PolicySetDefParam)
                }
                
                $PolicyAssignParam = $($args[2])
                # PolicyAssign can be PSCustomObject, Hashtable, or Array - preserve as-is
                
                $PolicyDefParam = $($args[3])
                if ($null -eq $PolicyDefParam) {
                    $PolicyDefParam = @()
                } elseif ($PolicyDefParam -isnot [System.Array]) {
                    $PolicyDefParam = @($PolicyDefParam)
                }

                $PolResult = Start-ARIPolicyJob -Subscriptions $SubsParam -PolicySetDef $PolicySetDefParam -PolicyAssign $PolicyAssignParam -PolicyDef $PolicyDefParam

                $PolResult

            } -ArgumentList $Subscriptions, $PolicySetDef, $PolicyAssign, $PolicyDef, $ARIModule | Out-Null
        }
    else
        {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Policy Job.')
            Start-Job -Name 'Policy' -ScriptBlock {

                import-module $($args[4])

                # Ensure variables are arrays after job serialization (arrays can become null or lose type)
                $SubsParam = $($args[0])
                if ($null -eq $SubsParam) {
                    $SubsParam = @()
                } elseif ($SubsParam -isnot [System.Array]) {
                    $SubsParam = @($SubsParam)
                }
                
                $PolicySetDefParam = $($args[1])
                if ($null -eq $PolicySetDefParam) {
                    $PolicySetDefParam = @()
                } elseif ($PolicySetDefParam -isnot [System.Array]) {
                    $PolicySetDefParam = @($PolicySetDefParam)
                }
                
                $PolicyAssignParam = $($args[2])
                # PolicyAssign can be PSCustomObject, Hashtable, or Array - preserve as-is
                
                $PolicyDefParam = $($args[3])
                if ($null -eq $PolicyDefParam) {
                    $PolicyDefParam = @()
                } elseif ($PolicyDefParam -isnot [System.Array]) {
                    $PolicyDefParam = @($PolicyDefParam)
                }

                $PolResult = Start-ARIPolicyJob -Subscriptions $SubsParam -PolicySetDef $PolicySetDefParam -PolicyAssign $PolicyAssignParam -PolicyDef $PolicyDefParam

                $PolResult

            } -ArgumentList $Subscriptions, $PolicySetDef, $PolicyAssign, $PolicyDef, $ARIModule | Out-Null
        }
}