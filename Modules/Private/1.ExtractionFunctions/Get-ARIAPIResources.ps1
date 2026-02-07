<#
.Synopsis
Module responsible for retrieving Azure API resources.

.DESCRIPTION
This module retrieves Azure API resources, including Resource Health, Managed Identities, Advisor Scores, and Policies.

.Link
https://github.com/microsoft/ARI/Modules/Private/1.ExtractionFunctions/Get-ARIAPIResources.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI).

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>
function Get-ARIAPIResources {
    Param($Subscriptions, $AzureEnvironment, $SkipPolicy)

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting API Inventory')

    try
        {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Acquiring Token')
            $Token = Get-AzAccessToken -AsSecureString -InformationAction SilentlyContinue -WarningAction SilentlyContinue -Debug:$false

            $TokenData = $Token.Token | ConvertFrom-SecureString -AsPlainText

            $header = @{
                'Authorization' = 'Bearer ' + $TokenData
            }
        }
    catch
        {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error: ' + $_.Exception.Message)
            return
        }
    

    if ($AzureEnvironment -eq 'AzureCloud') {
        $AzURL = 'management.azure.com'
    } 
    elseif ($AzureEnvironment -eq 'AzureUSGovernment') {
        $AzURL = 'management.usgovcloudapi.net'
    }
    elseif ($AzureEnvironment -eq 'AzureChinaCloud') {
        $AzURL = 'management.chinacloudapi.cn'
    }
    else {
        Write-Host ('Invalid Azure Environment for API Rest Inventory: ' + $AzureEnvironment) -ForegroundColor Red
        return
    }
    $ResourceHealthHistoryDate = (Get-Date).AddMonths(-6)
    $APIResults = @()

    function Get-ARIAllPages {
        param([string]$Url)

        $items = [System.Collections.ArrayList]::new()
        $nextUrl = $Url
        while (-not [string]::IsNullOrWhiteSpace($nextUrl)) {
            $resp = Invoke-RestMethod -Uri $nextUrl -Headers $header -Method GET
            if ($null -ne $resp) {
                if ($resp.PSObject.Properties['value'] -and $null -ne $resp.value) {
                    $pageItems = if ($resp.value -is [System.Array]) { $resp.value } else { @($resp.value) }
                    foreach ($item in $pageItems) { $null = $items.Add($item) }
                } else {
                    $null = $items.Add($resp)
                }
                $nextUrl = if ($resp.PSObject.Properties['nextLink']) { $resp.nextLink } else { $null }
            } else {
                $nextUrl = $null
            }
        }
        return $items.ToArray()
    }

    foreach ($Subscription in $Subscriptions)
        {
            $ResourceHealth = ""
            $Identities = ""
            $ADVScore = ""
            $ReservationRecon = ""
            $PolicyAssign = ""
            $PolicySetDef = ""
            $PolicyDef = ""

            $SubName = $Subscription.Name
            $Sub = $Subscription.id

            Write-Host 'Running API Inventory at: ' -NoNewline
            Write-Host $SubName -ForegroundColor Cyan

            #ResourceHealth Events
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Getting ResourceHealth Events')
            $url = ('https://' + $AzURL + '/subscriptions/' + $Sub + '/providers/Microsoft.ResourceHealth/events?api-version=2022-10-01&queryStartTime=' + $ResourceHealthHistoryDate)
            try {
                $ResourceHealth = Invoke-RestMethod -Uri $url -Headers $header -Method GET
            }
            catch {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error: ' + $_.Exception.Message)
                $ResourceHealth = ""
            }
            
            Start-Sleep -Milliseconds 200

            #Managed Identities
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Getting Managed Identities')
            $url = ('https://' + $AzURL + '/subscriptions/' + $Sub + '/providers/Microsoft.ManagedIdentity/userAssignedIdentities?api-version=2023-01-31')
            try {
                $Identities = Invoke-RestMethod -Uri $url -Headers $header -Method GET
            }
            catch {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error: ' + $_.Exception.Message)
                $Identities = ""
            }
            Start-Sleep -Milliseconds 200

            #Advisor Score
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Getting Advisor Score')
            $url = ('https://' + $AzURL + '/subscriptions/' + $Sub + '/providers/Microsoft.Advisor/advisorScore?api-version=2023-01-01')
            try {
                $ADVScore = Invoke-RestMethod -Uri $url -Headers $header -Method GET
            }
            catch {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error: ' + $_.Exception.Message)
                $ADVScore = ""
            }
            Start-Sleep -Milliseconds 200

            #VM Reservation Recommendation
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Getting VM Reservation Recommendation')
            $url = ('https://' + $AzURL + '/subscriptions/' + $Sub + '/providers/Microsoft.Consumption/reservationRecommendations?api-version=2023-05-01')
            try {
                $ReservationRecon = Invoke-RestMethod -Uri $url -Headers $header -Method GET
            }
            catch {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error: ' + $_.Exception.Message)
                $ReservationRecon = ""
            }
            Start-Sleep -Milliseconds 200

            # Handle SkipPolicy parameter - it might be a switch, boolean, or null
            $shouldSkipPolicy = $false
            if ($null -ne $SkipPolicy) {
                if ($SkipPolicy -is [switch]) {
                    $shouldSkipPolicy = $SkipPolicy.IsPresent
                } else {
                    # If it's not a switch, treat as boolean
                    $shouldSkipPolicy = [bool]$SkipPolicy
                }
            }
            
            if (!$shouldSkipPolicy)
                {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Getting Policies')
                    #Policies
                    try {
                        $url = ('https://'+ $AzURL +'/subscriptions/'+$sub+'/providers/Microsoft.PolicyInsights/policyStates/latest/summarize?api-version=2019-10-01')
                        $policySummarizeResponse = Invoke-RestMethod -Uri $url -Headers $header -Method POST
                        # Log summarize response shape for troubleshooting
                        $summaryType = $policySummarizeResponse.PSObject.TypeNames | Select-Object -First 1
                        $summaryKeys = $policySummarizeResponse.PSObject.Properties.Name -join ','
                        $summaryValue = $policySummarizeResponse.value
                        $summaryValueType = if ($null -ne $summaryValue) { $summaryValue.PSObject.TypeNames | Select-Object -First 1 } else { 'null' }
                        $summaryValueKeys = if ($null -ne $summaryValue -and ($summaryValue -is [PSCustomObject] -or $summaryValue -is [System.Collections.Hashtable])) { $summaryValue.PSObject.Properties.Name -join ',' } else { '' }
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Policy summarize response type=' + $summaryType + '; keys=' + $summaryKeys + '; valueType=' + $summaryValueType + '; valueKeys=' + $summaryValueKeys)

                        # Prefer .value, but fall back to top-level policyAssignments if present
                        if ($null -eq $summaryValue -and ($policySummarizeResponse.PSObject.Properties.Name -contains 'policyAssignments')) {
                            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Policy summarize response missing .value; using top-level policyAssignments')
                            $summaryValue = $policySummarizeResponse.policyAssignments
                        }

                        $PolicyAssign = $summaryValue
                        
                        # The summarize API returns an array in .value, where each element may have policyAssignments
                        # We need to extract all policyAssignments from all elements
                        $PolicyAssignCount = 0
                        $allPolicyAssignments = @()
                        
                        if ($null -ne $PolicyAssign) {
                            if ($PolicyAssign -is [System.Array]) {
                                # Iterate through array and collect all policyAssignments
                                foreach ($summaryItem in $PolicyAssign) {
                                    if ($null -ne $summaryItem) {
                                        $hasPolicyAssignments = $false
                                        if ($summaryItem -is [PSCustomObject]) {
                                            $hasPolicyAssignments = $summaryItem.PSObject.Properties.Name -contains 'policyAssignments'
                                        } elseif ($summaryItem -is [System.Collections.Hashtable]) {
                                            $hasPolicyAssignments = $summaryItem.ContainsKey('policyAssignments')
                                        }
                                        
                                        if ($hasPolicyAssignments -and $null -ne $summaryItem.policyAssignments) {
                                            if ($summaryItem.policyAssignments -is [System.Array]) {
                                                $allPolicyAssignments += $summaryItem.policyAssignments
                                            } elseif ($null -ne $summaryItem.policyAssignments) {
                                                $allPolicyAssignments += @($summaryItem.policyAssignments)
                                            }
                                        }
                                    }
                                }
                                $PolicyAssignCount = $allPolicyAssignments.Count
                                # Store the flattened array for later use
                                $PolicyAssign = if ($PolicyAssignCount -gt 0) { $allPolicyAssignments } else { @() }
                            } elseif ($PolicyAssign -is [PSCustomObject] -or $PolicyAssign -is [System.Collections.Hashtable]) {
                                # Single object - check for policyAssignments property
                                $hasPolicyAssignments = $false
                                if ($PolicyAssign -is [PSCustomObject]) {
                                    $hasPolicyAssignments = $PolicyAssign.PSObject.Properties.Name -contains 'policyAssignments'
                                } elseif ($PolicyAssign -is [System.Collections.Hashtable]) {
                                    $hasPolicyAssignments = $PolicyAssign.ContainsKey('policyAssignments')
                                }
                                
                                if ($hasPolicyAssignments -and $null -ne $PolicyAssign.policyAssignments) {
                                    if ($PolicyAssign.policyAssignments -is [System.Array]) {
                                        $PolicyAssignCount = $PolicyAssign.policyAssignments.Count
                                        $PolicyAssign = $PolicyAssign.policyAssignments
                                    } else {
                                        $PolicyAssignCount = 1
                                        $PolicyAssign = @($PolicyAssign.policyAssignments)
                                    }
                                } else {
                                    $PolicyAssignCount = 0
                                    $PolicyAssign = @()
                                }
                            } else {
                                $PolicyAssignCount = 0
                                $PolicyAssign = @()
                            }
                        } else {
                            $PolicyAssign = @()
                        }
                        
                        $policyAssignType = if ($null -ne $PolicyAssign) { $PolicyAssign.PSObject.TypeNames | Select-Object -First 1 } else { 'null' }
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Collected ' + $PolicyAssignCount + ' policy assignment(s) for subscription ' + $sub + '; PolicyAssignType=' + $policyAssignType)
                        Start-Sleep -Milliseconds 200
                        $url = ('https://'+ $AzURL +'/subscriptions/'+$sub+'/providers/Microsoft.Authorization/policySetDefinitions?api-version=2023-04-01')
                        $PolicySetDef = Get-ARIAllPages -Url $url
                        $PolicySetDefCount = if ($null -ne $PolicySetDef -and $PolicySetDef -is [System.Array]) { $PolicySetDef.Count } elseif ($null -ne $PolicySetDef) { 1 } else { 0 }
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Collected ' + $PolicySetDefCount + ' policy set definition(s)')
                        Start-Sleep -Milliseconds 200
                        $url = ('https://'+ $AzURL +'/subscriptions/'+$sub+'/providers/Microsoft.Authorization/policyDefinitions?api-version=2023-04-01')
                        $PolicyDef = Get-ARIAllPages -Url $url
                        $PolicyDefCount = if ($null -ne $PolicyDef -and $PolicyDef -is [System.Array]) { $PolicyDef.Count } elseif ($null -ne $PolicyDef) { 1 } else { 0 }
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Collected ' + $PolicyDefCount + ' policy definition(s)')
                    }
                    catch {
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Error collecting Policies: ' + $_.Exception.Message)
                        $PolicyAssign = ""
                        $PolicySetDef = ""
                        $PolicyDef = ""
                    }
                }

            Start-Sleep -Milliseconds 300

            $tmp = @{
                'Subscription'          = $Sub;
                'ResourceHealth'        = $ResourceHealth.value;
                'ManagedIdentities'     = $Identities.value;
                'AdvisorScore'          = $ADVScore.value;
                'ReservationRecomen'    = $ReservationRecon.value;
                'PolicyAssign'          = $PolicyAssign;
                'PolicyDef'             = $PolicyDef;
                'PolicySetDef'          = $PolicySetDef
            }
            $APIResults += $tmp

        }

        <#
        $Body = @{
            reportType = "OverallSummaryReport"
            subscriptionList = @($Subscri)
            carbonScopeList = @("Scope1")
            dateRange = @{
                start = "2024-06-01"
                end = "2024-06-30"
            }
        }
        $url = 'https://management.azure.com/providers/Microsoft.Carbon/carbonEmissionReports?api-version=2023-04-01-preview'
        #$url = 'https://management.azure.com/providers/Microsoft.Carbon/queryCarbonEmissionDataAvailableDateRange?api-version=2023-04-01-preview'

        $Carbon = Invoke-RestMethod -Uri $url -Headers $header -Body ($Body | ConvertTo-Json) -Method POST -ContentType 'application/json'

        

        $Today = Get-Date
        $EndDate = Get-Date -Year $Today.Year -Month $Today.Month -Day $Today.Day -Hour 23 -Minute 59 -Second 59 -Millisecond 0
        $Days = 60
        $StartDate = ($EndDate).AddDays(-$Days)

        $Hash = @{name="PreTaxCost";function="Sum"}
        $MHash = @{totalCost=$Hash}
        $Granularity = 'Monthly'

        $Grouping = @()
        $GTemp = @{Name='ResourceType';Type='Dimension'}
        $Grouping += $GTemp
        $GTemp = @{Name='ResourceGroup';Type='Dimension'}
        $Grouping += $GTemp

        $Body = @{
                type = "ActualCost"
                timeframe = "Custom"
                dataset = @{
                    granularity = $Granularity
                    aggregation = @($MHash)
                    }
                grouping = $Grouping
                timePeriod = @{
                    startDate = $StartDate
                    endDate = $EndDate
                }
        }

        $url = 'https://management.azure.com/subscriptions/$sub/providers/Microsoft.CostManagement/query?api-version=2019-11-01'

        $Cost = Invoke-RestMethod -Uri $url -Headers $header -Body ($Body | ConvertTo-Json -Depth 10) -Method POST -ContentType 'application/json'

        #>

        return $APIResults
}
