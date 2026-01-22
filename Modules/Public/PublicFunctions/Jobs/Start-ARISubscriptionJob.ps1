<#
.Synopsis
Start Subscription Job Module

.DESCRIPTION
This script processes and creates the Subscriptions sheet based on resources and their subscriptions.

.Link
https://github.com/microsoft/ARI/Modules/Public/PublicFunctions/Jobs/Start-ARISubscriptionJob.ps1

.COMPONENT
This powershell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola

#>
function Start-ARISubscriptionJob {
    param($Subscriptions, $Resources, $CostData)

    try {
        # Ensure Subscriptions is an array for safe .Count access
        if ($null -eq $Subscriptions) {
            $Subscriptions = @()
        } elseif ($Subscriptions -isnot [System.Array]) {
            $Subscriptions = @($Subscriptions)
        }

        # Debug: Log what we received
        $resourcesCount = if ($null -ne $Resources -and $Resources -is [System.Array]) { $Resources.Count } elseif ($null -ne $Resources) { 1 } else { 0 }
        $subscriptionsCount = if ($null -ne $Subscriptions -and $Subscriptions -is [System.Array]) { $Subscriptions.Count } elseif ($null -ne $Subscriptions) { 1 } else { 0 }
        Write-Debug "Start-ARISubscriptionJob: Received $resourcesCount resource(s), $subscriptionsCount subscription(s)"

    if ([string]::IsNullOrEmpty($CostData))
        {
            # Ensure Resources is an array
            if ($null -eq $Resources) {
                $Resources = @()
            }
            if ($Resources -isnot [System.Array]) {
                $Resources = @($Resources)
            }
            
            Write-Debug "Start-ARISubscriptionJob: Resources array has $($Resources.Count) item(s)"
            
            # Filter resources - handle both lowercase 'type' and uppercase 'Type'
            $ResTable = $Resources | Where-Object { 
                $resourceType = if ($null -ne $_.type) { $_.type } elseif ($null -ne $_.Type) { $_.Type } else { '' }
                $resourceType -notin ('microsoft.advisor/recommendations',
                                    'ARI/VM/Quotas',
                                    'ARI/VM/SKU',
                                    'Microsoft.Advisor/advisorScore',
                                    'Microsoft.ResourceHealth/events',
                                    'microsoft.support/supporttickets' )
            }
            
            # Select properties - handle case variations
            $resTable2 = $ResTable | Select-Object @{
                Name='id'; Expression={if ($null -ne $_.id) { $_.id } elseif ($null -ne $_.ID) { $_.ID } else { $_.Id }}
            }, @{
                Name='Type'; Expression={if ($null -ne $_.type) { $_.type } elseif ($null -ne $_.Type) { $_.Type } else { '' }}
            }, @{
                Name='location'; Expression={if ($null -ne $_.location) { $_.location } elseif ($null -ne $_.Location) { $_.Location } else { '' }}
            }, @{
                Name='resourcegroup'; Expression={if ($null -ne $_.resourcegroup) { $_.resourcegroup } elseif ($null -ne $_.resourceGroup) { $_.resourceGroup } elseif ($null -ne $_.ResourceGroup) { $_.ResourceGroup } elseif ($null -ne $_.'Resource Group') { $_.'Resource Group' } else { '' }}
            }, @{
                Name='subscriptionid'; Expression={if ($null -ne $_.subscriptionid) { $_.subscriptionid } elseif ($null -ne $_.subscriptionId) { $_.subscriptionId } elseif ($null -ne $_.SubscriptionId) { $_.SubscriptionId } else { '' }}
            }
            
            Write-Debug "Start-ARISubscriptionJob: After filtering, ResTable has $($ResTable.Count) item(s)"
            Write-Debug "Start-ARISubscriptionJob: resTable2 has $($resTable2.Count) item(s)"
            
            $ResTable3 = $resTable2 | Group-Object -Property Type, location, resourcegroup, subscriptionid
            
            # Safely get ResTable3 count
            $resTable3Count = if ($null -ne $ResTable3 -and $ResTable3 -is [System.Array]) { $ResTable3.Count } elseif ($null -ne $ResTable3) { 1 } else { 0 }
            Write-Debug "Start-ARISubscriptionJob: After grouping, ResTable3 has $resTable3Count group(s)"

            # Initialize FormattedTable as empty array to ensure it's always an array
            $FormattedTable = @()
            if ($null -ne $ResTable3) {
                # Ensure ResTable3 is an array (Group-Object can return single object)
                $ResTable3Array = if ($ResTable3 -is [System.Array]) { $ResTable3 } else { @($ResTable3) }
                
                $FormattedTable = foreach ($ResourcesSUB in $ResTable3Array) 
                    {
                        # Safely get ResourcesSUB.Count - Group-Object results have Count property
                        $resourcesSubCount = 0
                        if ($null -ne $ResourcesSUB) {
                            if ($ResourcesSUB -is [Microsoft.PowerShell.Commands.GroupInfo]) {
                                $resourcesSubCount = $ResourcesSUB.Count
                            } elseif ($ResourcesSUB -is [System.Collections.ICollection]) {
                                $resourcesSubCount = $ResourcesSUB.Count
                            } else {
                                # Try to access Count property safely
                                try {
                                    $resourcesSubCount = $ResourcesSUB.Count
                                } catch {
                                    $resourcesSubCount = 1  # Default to 1 if Count not available
                                }
                            }
                        }
                        
                        $ResourceDetails = $ResourcesSUB.name -split ", "
                        # Ensure ResourceDetails is an array
                        if ($null -eq $ResourceDetails) {
                            $ResourceDetails = @()
                        } elseif ($ResourceDetails -isnot [System.Array]) {
                            $ResourceDetails = @($ResourceDetails)
                        }
                        
                        if ($ResourceDetails.Count -ge 4) {
                            $subId = $ResourceDetails[3]
                            $SubName = $Subscriptions | Where-Object { 
                                $subObjId = if ($null -ne $_.Id) { $_.Id } elseif ($null -ne $_.id) { $_.id } elseif ($null -ne $_.ID) { $_.ID } else { '' }
                                $subObjId -eq $subId
                            } | Select-Object -First 1
                            
                            $subscriptionName = if ($null -ne $SubName -and $null -ne $SubName.Name) { $SubName.Name } elseif ($null -ne $SubName -and $null -ne $SubName.name) { $SubName.name } else { $subId }
                            
                            $obj = [PSCustomObject]@{
                                'Subscription'      = $subscriptionName
                                'SubscriptionId'    = $subId
                                'Resource Group'    = if ($ResourceDetails.Count -ge 3) { $ResourceDetails[2] } else { '' }
                                'Location'          = if ($ResourceDetails.Count -ge 2) { $ResourceDetails[1] } else { '' }
                                'Resource Type'     = if ($ResourceDetails.Count -ge 1) { $ResourceDetails[0] } else { '' }
                                'Resources Count'   = $resourcesSubCount
                            }
                            $obj
                        }
                    }
                # Ensure FormattedTable is an array (foreach might return null if empty)
                if ($null -eq $FormattedTable) {
                    $FormattedTable = @()
                } elseif ($FormattedTable -isnot [System.Array]) {
                    $FormattedTable = @($FormattedTable)
                }
            }
            
            Write-Debug "Start-ARISubscriptionJob: FormattedTable has $($FormattedTable.Count) item(s)"
        }
    else
        {
            # Initialize FormattedTable as empty array to ensure it's always an array
            $FormattedTable = @()
            if ($null -ne $CostData) {
                $FormattedTable = foreach ($Cost in $CostData)
                    {
                        Foreach ($CostDetail in $Cost.CostData.Row)
                            {
                                Foreach ($Currency in $CostDetail[6])
                                    {
                                        $Date0 = [datetime]$CostDetail[1]
                                        $DateMonth = ((Get-Culture).DateTimeFormat.GetMonthName(([datetime]$Date0).ToString("MM"))).ToString()
                                        $DateYear = (([datetime]$Date0).ToString("yyyy")).ToString()

                                        $obj = [PSCustomObject]@{
                                            'Subscription'      = $Cost.SubscriptionName
                                            'SubscriptionId'    = $Cost.SubscriptionId
                                            'Resource Group'    = $CostDetail[3]
                                            'Resource Type'     = $CostDetail[2]
                                            'Location'          = $CostDetail[4]
                                            'Service Name'      = $CostDetail[5]
                                            'Currency'          = $Currency
                                            'Cost'              = $CostDetail[0]
                                            'Detailed Cost'     = $CostDetail[0]
                                            'Year'              = $DateYear
                                            'Month'             = $DateMonth
                                        }
                                        $obj
                                    }
                            }
                    }
                # Ensure FormattedTable is an array (foreach might return null if empty)
                if ($null -eq $FormattedTable) {
                    $FormattedTable = @()
                } elseif ($FormattedTable -isnot [System.Array]) {
                    $FormattedTable = @($FormattedTable)
                }
            }
        }

        <#
        $outerKeyGeneral = [Func[Object,string]] { $args[0].SubscriptionID, $args[0].ResourceGroup, $args[0].ResourceType }
        $innerKeyGeneral = [Func[Object,string]] { $args[0].SubscriptionID, $args[0].ResourceGroup, $args[0].ResourceType }

        $ResultDelegate = [Func[Object, Object, PSCustomObject]] {
            param
            (
                $SubTable,
                $CostTable
            )
            [PSCustomObject]@{
                'Subscription' = $SubTable.Subscription
                'Resource Group' = $SubTable.ResourceGroup
                'Location' = $CostTable.Location
                'Resource Type' = $SubTable.ResourceType
                'Resources Count' = $SubTable.ResourcesCount
                'Currency' = $CostTable.Currency
                'Cost' = $CostTable.Cost
                'Detailed Cost' = $CostTable.DetailedCost
                'Year' = $CostTable.Year
                'Month' = $CostTable.Month
            }  
        }

        [System.Func[System.Object, [Collections.Generic.IEnumerable[System.Object]], System.Object]]$query = {
            param(
                $SubTable,
                $CostTable
            )
            $RightJoin = [System.Linq.Enumerable]::SingleOrDefault($CostTable)

            [PSCustomObject]@{
                'Subscription' = $SubTable.Subscription
                'Resource Group' = $SubTable.ResourceGroup
                'Location' = $SubTable.Location
                'Resource Type' = $SubTable.ResourceType
                'Resources Count' = $SubTable.ResourcesCount
                'Currency' = $RightJoin.Currency
                'Cost' = $RightJoin.Cost
                'Year' = $RightJoin.Year
                'Month' = $RightJoin.Month
            }
        }

        $LeftJoin = [System.Linq.Enumerable]::ToArray([System.Linq.Enumerable]::GroupJoin($SubDetailsTable, $CostDetailsTable, $outerKeyGroup, $innerKeyGeneraltest, $query))


        $InnerJoinResult = [System.Linq.Enumerable]::ToArray([System.Linq.Enumerable]::Join($SubDetailsTable, $CostDetailsTable, $outerKeyGeneral, $innerKeyGeneral, $resultDelegate))

        $InnerJoinResult
        #>

        $FormattedTable
    } catch {
        Write-Error "Error in Start-ARISubscriptionJob: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        Write-Error "Line: $($_.InvocationInfo.ScriptLineNumber)"
        Write-Error "Function: $($_.InvocationInfo.FunctionName)"
        Write-Error "Subscriptions count: $($Subscriptions.Count)"
        Write-Error "Resources type: $($Resources.GetType().FullName)"
        throw
    }
}
