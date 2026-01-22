<#
.Synopsis
Module responsible for coordinate the extraction of Resource and build the Graph queries

.DESCRIPTION
This module is the main module for the Azure Resource Graphs that will be run against the environment.

.Link
https://github.com/microsoft/ARI/Modules/Private/1.ExtractionFunctions/Start-ARIGraphExtraction.ps1

.COMPONENT
This powershell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.11
First Release Date: 15th Oct, 2024
Authors: Claudio Merola

#>
Function Start-ARIGraphExtraction {
    Param($ManagementGroup, $Subscriptions, $SubscriptionID, $ResourceGroup, $SecurityCenter, $SkipAdvisory, $IncludeTags, $TagKey, $TagValue, $AzureEnvironment)

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Extractor function')

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Powershell Edition: ' + ([string]$psversiontable.psEdition))
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Powershell Version: ' + ([string]$psversiontable.psVersion))

    #Field for tags
    if ($IncludeTags.IsPresent) {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+"Tags will be included")
        $GraphQueryTags = ",tags "
    } else {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+"Tags will be ignored")
        $GraphQueryTags = ""
    }

    <###################################################### Subscriptions ######################################################################>

    Write-Progress -activity 'Azure Inventory' -Status "2% Complete." -PercentComplete 2 -CurrentOperation 'Discovering Subscriptions..'

    if (![string]::IsNullOrEmpty($ManagementGroup))
        {
            $Subscriptions = Get-ARIManagementGroups -ManagementGroup $ManagementGroup -Subscriptions $Subscriptions
        }

    # Safely access Subscriptions.id.count - handle null/empty cases
    if ($null -ne $Subscriptions -and $null -ne $Subscriptions.id) {
        $SubCount = [string]$Subscriptions.id.count
    } elseif ($null -ne $Subscriptions -and $Subscriptions -is [System.Array]) {
        $SubCount = [string]$Subscriptions.Count
    } else {
        $SubCount = "0"
    }

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Number of Subscriptions Found: ' + $SubCount)
    Write-Progress -activity 'Azure Inventory' -Status "3% Complete." -PercentComplete 3 -CurrentOperation "$SubCount Subscriptions found.."

    <######################################################## INVENTORY LOOPs #######################################################################>

    Write-Progress -activity 'Azure Inventory' -Status "4% Complete." -PercentComplete 4 -CurrentOperation "Starting Resources extraction jobs.."

    if(![string]::IsNullOrEmpty($ResourceGroup) -and [string]::IsNullOrEmpty($SubscriptionID))
        {
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Resource Group Name present, but missing Subscription ID.')
            Write-Output ''
            Write-Output 'If Using the -ResourceGroup Parameter, the Subscription ID must be informed'
            Write-Output ''
            Exit
        }
    else
        {
            # Safely access Subscriptions.id - handle null/empty cases
            if ($null -ne $Subscriptions -and $null -ne $Subscriptions.id) {
                $Subscri = $Subscriptions.id
            } elseif ($null -ne $Subscriptions -and $Subscriptions -is [System.Array]) {
                # If Subscriptions is an array, extract id property from each
                $Subscri = $Subscriptions | ForEach-Object { if ($null -ne $_.id) { $_.id } } | Where-Object { $_ -ne $null }
            } else {
                $Subscri = @()
            }
            $RGQueryExtension = ''
            $TagQueryExtension = ''
            $MGQueryExtension = ''
            if(![string]::IsNullOrEmpty($ResourceGroup) -and ![string]::IsNullOrEmpty($SubscriptionID))
                {
                    $RGQueryExtension = "| where resourceGroup in~ ('$([String]::Join("','",$ResourceGroup))')"
                }
            elseif(![string]::IsNullOrEmpty($TagKey) -or ![string]::IsNullOrEmpty($TagValue))
                {

                    $TagQueryExtension = "| where isnotempty(tags) | mvexpand tags | extend tagKey = tostring(bag_keys(tags)[0]) | extend tagValue = tostring(tags[tagKey]) "

                    if (![string]::IsNullOrEmpty($TagKey)){ 
                        $TagQueryExtension = $TagQueryExtension + "| where tagKey =~ '$TagKey'"
                    }

                    if (![string]::IsNullOrEmpty($TagValue)){ 
                        $TagQueryExtension = $TagQueryExtension + " and tagValue =~ '$TagValue'"
                    }

                    #$TagQueryExtension = "| where isnotempty(tags) | mvexpand tags | extend tagKey = tostring(bag_keys(tags)[0]) | extend tagValue = tostring(tags[tagKey]) | where tagKey =~ '$TagKey' and tagValue =~ '$TagValue'"
                }
            elseif (![string]::IsNullOrEmpty($ManagementGroup))
                {
                    $MGQueryExtension = "| join kind=inner (resourcecontainers | where type == 'microsoft.resources/subscriptions' | mv-expand managementGroupParent = properties.managementGroupAncestorsChain | where managementGroupParent.name =~ '$ManagementGroup' | project subscriptionId, managanagementGroup = managementGroupParent.name) on subscriptionId"
                    $MGContainerExtension = "| mv-expand managementGroupParent = properties.managementGroupAncestorsChain | where managementGroupParent.name =~ '$ManagementGroup'"
                }
        }

            $ExcludedTypes = "| where type !in ('microsoft.logic/workflows','microsoft.portal/dashboards','microsoft.resources/templatespecs/versions','microsoft.resources/templatespecs')"

            # Initialize Resources array if not already initialized
            if ($null -eq $Resources) {
                $Resources = @()
            }

            $GraphQuery = "resources $RGQueryExtension $TagQueryExtension $MGQueryExtension $ExcludedTypes | project id,name,type,tenantId,kind,location,resourceGroup,subscriptionId,managedBy,sku,plan,properties,identity,zones,extendedLocation$($GraphQueryTags) | order by id asc"

            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Invoking Inventory Loop for Resources')
            $loopResult = Invoke-ARIInventoryLoop -GraphQuery $GraphQuery -FSubscri $Subscri -LoopName 'Resources'
            if ($null -ne $loopResult) {
                $Resources += $loopResult
            }

            $GraphQuery = "networkresources $RGQueryExtension $TagQueryExtension $MGQueryExtension | project id,name,type,tenantId,kind,location,resourceGroup,subscriptionId,managedBy,sku,plan,properties,identity,zones,extendedLocation$($GraphQueryTags) | order by id asc"

            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Invoking Inventory Loop for Network Resources')
            $loopResult = Invoke-ARIInventoryLoop -GraphQuery $GraphQuery -FSubscri $Subscri -LoopName 'Network Resources'
            if ($null -ne $loopResult) {
                $Resources += $loopResult
            }

            if ($AzureEnvironment -ne 'AzureUSGovernment')
                {
                    $GraphQuery = "SupportResources $RGQueryExtension $TagQueryExtension $MGQueryExtension | project id,name,type,tenantId,kind,location,resourceGroup,subscriptionId,managedBy,sku,plan,properties,identity,zones,extendedLocation$($GraphQueryTags) | order by id asc"

                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Invoking Inventory Loop for Support Tickets')
                    $loopResult = Invoke-ARIInventoryLoop -GraphQuery $GraphQuery -FSubscri $Subscri -LoopName 'SupportTickets'
                    if ($null -ne $loopResult) {
                        $Resources += $loopResult
                    }
                }

            $GraphQuery = "recoveryservicesresources $RGQueryExtension $TagQueryExtension | where type =~ 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems' or type =~ 'microsoft.recoveryservices/vaults/backuppolicies' $MGQueryExtension  | project id,name,type,tenantId,kind,location,resourceGroup,subscriptionId,managedBy,sku,plan,properties,identity,zones,extendedLocation$($GraphQueryTags) | order by id asc"

            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Invoking Inventory Loop for Backup Resources')
            $loopResult = Invoke-ARIInventoryLoop -GraphQuery $GraphQuery -FSubscri $Subscri -LoopName 'Backup Items'
            if ($null -ne $loopResult) {
                $Resources += $loopResult
            }

            $GraphQuery = "desktopvirtualizationresources $RGQueryExtension $MGQueryExtension| project id,name,type,tenantId,kind,location,resourceGroup,subscriptionId,managedBy,sku,plan,properties,identity,zones,extendedLocation$($GraphQueryTags) | order by id asc"

            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Invoking Inventory Loop for AVD Resources')
            $loopResult = Invoke-ARIInventoryLoop -GraphQuery $GraphQuery -FSubscri $Subscri -LoopName 'Virtual Desktop'
            if ($null -ne $loopResult) {
                $Resources += $loopResult
            }

            $GraphQuery = "resourcecontainers $RGQueryExtension $TagQueryExtension $MGContainerExtension | project id,name,type,tenantId,kind,location,resourceGroup,subscriptionId,managedBy,sku,plan,properties,identity,zones,extendedLocation$($GraphQueryTags) | order by id asc"

            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Invoking Inventory Loop for Resource Containers')
            $ResourceContainers = Invoke-ARIInventoryLoop -GraphQuery $GraphQuery -FSubscri $Subscri -LoopName 'Subscriptions and Resource Groups'

            # Safely access ResourceContainers.count - handle null/empty cases
            if ($null -ne $ResourceContainers) {
                $ContainerCount = $ResourceContainers.count
            } else {
                $ContainerCount = 0
                $ResourceContainers = @()
            }
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Number of Resource Containers: '+ $ContainerCount)

            if (!($SkipAdvisory.IsPresent))
                {
                    $GraphQuery = "advisorresources $RGQueryExtension $MGQueryExtension | where properties.impact in~ ('Medium','High') | order by id asc"

                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Invoking Inventory Loop for Advisories')
                    $Advisories = Invoke-ARIInventoryLoop -GraphQuery $GraphQuery -FSubscri $Subscri -LoopName 'Advisories'

                    # Safely access Advisories.count - handle null/empty cases
                    if ($null -ne $Advisories) {
                        $AdvisorCount = $Advisories.count
                    } else {
                        $AdvisorCount = 0
                        $Advisories = @()
                    }
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Number of Advisors: '+ $AdvisorCount)
                }
            if ($SecurityCenter.IsPresent)
                {
                    $GraphQuery = "securityresources $RGQueryExtension | where type =~ 'microsoft.security/assessments' and properties['status']['code'] == 'Unhealthy' $MGQueryExtension | order by id asc"

                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Invoking Inventory Loop for Security Resources')
                    $Security = Invoke-ARIInventoryLoop -GraphQuery $GraphQuery -FSubscri $Subscri -LoopName 'Security Center'

                    # Safely access Security.count - handle null/empty cases
                    if ($null -ne $Security) {
                        $SecurityCount = $Security.count
                    } else {
                        $SecurityCount = 0
                        $Security = @()
                    }
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Number of Security Center Advisors: '+ $SecurityCount)
                } else {
                    # Initialize Security as empty array if SecurityCenter is not present
                    $Security = @()
                }

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Invoking Inventory Loop for Retirements')

    $RootPath = (get-item $PSScriptRoot).parent

    # Build path correctly - Join-Path only accepts two arguments, so chain them
    $RetirementPath = Join-Path $RootPath '3.ReportingFunctions'
    $RetirementPath = Join-Path $RetirementPath 'StyleFunctions'
    $RetirementPath = Join-Path $RetirementPath 'Retirement.kql'

    # Check if file exists, if not skip retirement query
    if (Test-Path $RetirementPath) {
        $RetirementQuery = Get-Content -Path $RetirementPath | Out-String
    } else {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Retirement.kql not found, skipping retirement query')
        $RetirementQuery = ""
    }

    # Only invoke retirement loop if query exists
    if (![string]::IsNullOrEmpty($RetirementQuery)) {
        $ResourceRetirements = Invoke-ARIInventoryLoop -GraphQuery $RetirementQuery -FSubscri $Subscri -LoopName 'Retirements'
    } else {
        $ResourceRetirements = @()
    }

    # Safely access ResourceRetirements.count - handle null/empty cases
    if ($null -ne $ResourceRetirements) {
        $RetirementCount = $ResourceRetirements.count
    } else {
        $RetirementCount = 0
        $ResourceRetirements = @()
    }

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Number of Retirements: '+ $RetirementCount)

    Write-Progress -activity 'Azure Inventory' -PercentComplete 10

    # Ensure all return values are arrays (not null) to prevent .count errors
    if ($null -eq $Resources) { $Resources = @() }
    if ($null -eq $ResourceContainers) { $ResourceContainers = @() }
    if ($null -eq $Advisories) { $Advisories = @() }
    if ($null -eq $Security) { $Security = @() }
    if ($null -eq $ResourceRetirements) { $ResourceRetirements = @() }

    $tmp = [PSCustomObject]@{
        Resources              = $Resources
        ResourceContainers     = $ResourceContainers
        Advisories             = $Advisories
        Security               = $Security
        Retirements            = $ResourceRetirements
    }
    return $tmp
}