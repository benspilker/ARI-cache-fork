<#
.SYNOPSIS
    This script creates Excel file to Analyze Azure Resources inside a Tenant

.DESCRIPTION
    Do you want to analyze your Azure Advisories in a table format? Document it in xlsx format.

.PARAMETER TenantID
    Specify the tenant ID you want to create a Resource Inventory.

    >>> IMPORTANT: YOU NEED TO USE THIS PARAMETER FOR TENANTS WITH MULTI-FACTOR AUTHENTICATION. <<<

.PARAMETER SubscriptionID
    Use this parameter to collect a specific Subscription in a Tenant

.PARAMETER ManagementGroup
    Use this parameter to collect all Subscriptions in a Specific Management Group in a Tenant

.PARAMETER Lite
    Use this parameter to use only the Import-Excel module and don't create the charts (using Excel's API)

.PARAMETER SecurityCenter
    Use this parameter to collect Security Center Advisories

.PARAMETER SkipAdvisory
    Use this parameter to skip the capture of Azure Advisories

.PARAMETER SkipPolicy
    Use this parameter to skip the capture of Azure Policies

.PARAMETER QuotaUsage
    Use this parameter to include Quota information

.PARAMETER IncludeTags
    Use this parameter to include Tags of every Azure Resources

.PARAMETER Debug
    Output detailed debug information.

.PARAMETER AzureEnvironment
    Specifies the Azure Cloud Environment to use. Default is 'AzureCloud'.

.PARAMETER Overview
    Specifies the Excel overview sheet design. Each value will change the main charts in the Overview sheet. Valid values are 1, 2, or 3. Default is 1.

.PARAMETER AppId
    Specifies the Application ID used to connect to Azure as a service principal. Requires TenantID and Secret.

.PARAMETER Secret
    Specifies the Secret used with the Application ID to connect to Azure as a service principal. Requires TenantID and AppId.

.PARAMETER CertificatePath
    Specifies the Certificate path used with the Application ID to connect to Azure as a service principal. Requires TenantID, AppId, and Secret.

.PARAMETER ResourceGroup
    Specifies one or more unique Resource Groups to be inventoried. Requires SubscriptionID.

.PARAMETER TagKey
    Specifies the tag key to be inventoried. Requires SubscriptionID.

.PARAMETER TagValue
    Specifies the tag value to be inventoried. Requires SubscriptionID.

.PARAMETER Heavy
    Use this parameter to enable heavy mode. This will force the job's load to be split into smaller batches. Avoiding CPU overload.

.PARAMETER NoAutoUpdate
    Use this parameter to skip automatic module updates.

.PARAMETER SkipAPIs
    Use this parameter to skip the capture of resources trough REST API.

.PARAMETER Automation
    Use this parameter to run in automation mode.

.PARAMETER StorageAccount
    Specifies the Storage Account name for storing the report.

.PARAMETER StorageContainer
    Specifies the Storage Container name for storing the report.

.PARAMETER Help
    Use this parameter to display the help information.

.PARAMETER DeviceLogin
    Use this parameter to enable device login.

.PARAMETER DiagramFullEnvironment
    Use this parameter to include the full environment in the diagram. By default the Network Topology Diagram will only include VNETs that are peered with other VNETs, this parameter will force the diagram to include all VNETs.

.PARAMETER ReportName
    Specifies the name of the report. Default is 'AzureResourceInventory'.

.PARAMETER ReportDir
    Specifies the directory where the report will be saved.

.EXAMPLE
    Default utilization. Read all tenants you have privileges, select a tenant in menu and collect from all subscriptions:
    PS C:\> Invoke-ARI

    Define the Tenant ID:
    PS C:\> Invoke-ARI -TenantID <your-Tenant-Id>

    Define the Tenant ID and for a specific Subscription:
    PS C:\> Invoke-ARI -TenantID <your-Tenant-Id> -SubscriptionID <your-Subscription-Id>

.NOTES
    AUTHORS: Claudio Merola and Renato Gregio | Azure Infrastucture/Automation/Devops/Governance

    Copyright (c) 2018 Microsoft Corporation. All rights reserved.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
    THE SOFTWARE.

.LINK
    Official Repository: https://github.com/microsoft/ARI
#>
Function Invoke-CachedARI-Patched {
    [CmdletBinding(PositionalBinding=$false)]
    param (
        [ValidateSet(1, 2, 3)]
        [int]$Overview = 1,    
        [ValidateSet('AzureCloud', 'AzureUSGovernment', 'AzureChinaCloud', 'AzureGermanCloud')]
        [string]$AzureEnvironment = 'AzureCloud',
        [string]$TenantID,
        [string]$AppId,
        [string]$Secret,
        [string]$CertificatePath,
        [string]$ReportName = 'AzureResourceInventory',
        [string]$ReportDir,
        [string]$StorageAccount,
        [string]$StorageContainer,
        [String[]]$SubscriptionID,
        [string[]]$ManagementGroup,
        [string[]]$ResourceGroup,
        [string[]]$TagKey,
        [string[]]$TagValue,
        [switch]$SecurityCenter,
        [switch]$Heavy,
        [Alias("SkipAdvisories","NoAdvisory","SkipAdvisor")]
        [switch]$SkipAdvisory,
        [Alias("DisableAutoUpdate","SkipAutoUpdate")]
        [switch]$NoAutoUpdate,
        [Alias("NoPolicy","SkipPolicies")]
        [switch]$SkipPolicy,
        [Alias("NoAPI","SkipAPI")]
        [switch]$SkipAPIs,
        [Alias("IncludeTag","AddTags")]
        [switch]$IncludeTags,
        [Alias("SkipVMDetail","NoVMDetails")]
        [switch]$SkipVMDetails,
        [Alias("Costs","IncludeCost")]
        [switch]$IncludeCosts,
        [switch]$QuotaUsage,
        [switch]$SkipDiagram,
        [switch]$Automation,
        [Alias("Low","Light")]
        [switch]$Lite,
        [switch]$Help,
        [switch]$DeviceLogin,
        [switch]$DiagramFullEnvironment,
        [switch]$UseExistingCache,
        [Alias("NoExcel","SkipReport")]
        [switch]$SkipExcel
        )

    # DEBUG: Log switch parameter state immediately
    Write-Host "[DEBUG] UseExistingCache parameter check:" -ForegroundColor Magenta
    Write-Host "  UseExistingCache.IsPresent: $($UseExistingCache.IsPresent)" -ForegroundColor Magenta
    Write-Host "  UseExistingCache value: $UseExistingCache" -ForegroundColor Magenta
    Write-Host "  PSBoundParameters contains UseExistingCache: $($PSBoundParameters.ContainsKey('UseExistingCache'))" -ForegroundColor Magenta
    if ($PSBoundParameters.ContainsKey('UseExistingCache')) {
        Write-Host "  PSBoundParameters['UseExistingCache']: $($PSBoundParameters['UseExistingCache'])" -ForegroundColor Magenta
    }

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Debugging Mode: On. ErrorActionPreference was set to "Continue", every error will be presented.')

    if ($DebugPreference -eq 'SilentlyContinue')
        {
            Write-Host 'Debugging Mode: ' -nonewline
            Write-Host 'Off' -ForegroundColor Yellow
            Write-Host 'Use the parameter ' -nonewline
            Write-Host '-Debug' -nonewline -ForegroundColor Yellow
            Write-Host ' to see debugging information during the inventory execution.'
            Write-Host 'For large environments, it is recommended to use the -Debug parameter to monitor the progress.' -ForegroundColor Yellow
        }

    if ($IncludeTags.IsPresent) { $InTag = $true } else { $InTag = $false }

    if ($Lite.IsPresent) { $RunLite = $true }else { $RunLite = $false }
    if ($DiagramFullEnvironment.IsPresent) {$FullEnv = $true}else{$FullEnv = $false}
    if ($Automation.IsPresent) 
        {
            $SkipAPIs = $true
            $RunLite = $true
            if (!$StorageAccount -or !$StorageContainer)
                {
                    Write-Output "Storage Account and Container are required for Automation mode. Aborting."
                    exit
                }
        }
    if ($Overview -eq 1 -and $SkipAPIs)
        {
            $Overview = 2
        }
    $TableStyle = "Light19"

    <#########################################################          Help          ######################################################################>

    Function Get-ARIUsageMode() {
        Write-Host ""
        Write-Host "Parameters"
        Write-Host ""
        Write-Host " -TenantID <ID>           :  Specifies the Tenant to be inventoried. "
        Write-Host " -SubscriptionID <ID>     :  Specifies Subscription(s) to be inventoried. "
        Write-Host " -ResourceGroup <NAME>    :  Specifies one (or more) unique Resource Group to be inventoried, This parameter requires the -SubscriptionID to work. "
        Write-Host " -AppId <ID>              :  Specifies the ApplicationID that is used to connect to Azure as service principal. This parameter requires the -TenantID and -Secret to work. "
        Write-Host " -Secret <VALUE>          :  Specifies the Secret that is used with the Application ID to connect to Azure as service principal. This parameter requires the -TenantID and -AppId to work. If -CertificatePath is also used the Secret value should be the Certifcate password instead of the Application secret. "
        Write-Host " -CertificatePath <PATH>  :  Specifies the Certificate path that is used with the Application ID to connect to Azure as service principal. This parameter requires the -TenantID, -AppId and -Secret to work. The required certificate format is pkcs#12. "
        Write-Host " -TagKey <NAME>           :  Specifies the tag key to be inventoried, This parameter requires the -SubscriptionID to work. "
        Write-Host " -TagValue <NAME>         :  Specifies the tag value be inventoried, This parameter requires the -SubscriptionID to work. "
        Write-Host " -SkipAdvisory            :  Do not collect Azure Advisory. "
        Write-Host " -SkipPolicy              :  Do not collect Azure Policies. "
        Write-Host " -SecurityCenter          :  Include Security Center Data. "
        Write-Host " -IncludeTags             :  Include Resource Tags. "
        Write-Host " -Online                  :  Use Online Modules. "
        Write-Host " -Debug                   :  Run in a Debug mode. "
        Write-Host " -AzureEnvironment        :  Change the Azure Cloud Environment. "
        Write-Host " -ReportName              :  Change the Default Name of the report. "
        Write-Host " -ReportDir               :  Change the Default Path of the report. "
        Write-Host ""
        Write-Host ""
        Write-Host ""
        Write-Host "Usage Mode and Examples: "
        Write-Host "If you do not specify Resource Inventory will be performed on all subscriptions for the selected tenant. "
        Write-Host "e.g. /> Invoke-ARI"
        Write-Host ""
        Write-Host "To perform the inventory in a specific Tenant and subscription use <-TenantID> and <-SubscriptionID> parameter "
        Write-Host "e.g. /> Invoke-ARI -TenantID <Azure Tenant ID> -SubscriptionID <Subscription ID>"
        Write-Host ""
        Write-Host "Including Tags:"
        Write-Host " By Default Azure Resource inventory do not include Resource Tags."
        Write-Host " To include Tags at the inventory use <-IncludeTags> parameter. "
        Write-Host "e.g. /> Invoke-ARI -TenantID <Azure Tenant ID> -IncludeTags"
        Write-Host ""
        Write-Host "Skipping Azure Advisor:"
        Write-Host " By Default Azure Resource inventory collects Azure Advisor Data."
        Write-Host " To ignore this  use <-SkipAdvisory> parameter. "
        Write-Host "e.g. /> Invoke-ARI -TenantID <Azure Tenant ID> -SubscriptionID <Subscription ID> -SkipAdvisory"
        Write-Host ""
        Write-Host "Using the latest modules :"
        Write-Host " You can use the latest modules. For this use <-Online> parameter."
        Write-Host " It's a pre-requisite to have internet access for ARI GitHub repo"
        Write-Host "e.g. /> Invoke-ARI -TenantID <Azure Tenant ID> -Online"
        Write-Host ""
        Write-Host "Running in Debug Mode :"
        Write-Host " To run in a Debug Mode use <-Debug> parameter."
        Write-Host ".e.g. /> Invoke-ARI -TenantID <Azure Tenant ID> -Debug"
        Write-Host ""
    }

    $TotalRunTime = [System.Diagnostics.Stopwatch]::StartNew()

    if ($Help.IsPresent) {
        Get-ARIUsageMode
        Exit
    }

    # Skip authentication and extraction when using existing cache
    # Force NoAutoUpdate when using existing cache to prevent breaking the patched version
    # Check both IsPresent and direct boolean check for switch parameter
    $useCache = $UseExistingCache.IsPresent -or $UseExistingCache
    if ($useCache) {
        # CRITICAL: Force NoAutoUpdate switch to TRUE to prevent auto-update from breaking patched version
        # For switch parameters, we need to add it to PSBoundParameters or set the variable directly
        $PSBoundParameters['NoAutoUpdate'] = $true
        $NoAutoUpdate = [switch]$true
        Write-Host "[UseExistingCache] Skipping resource extraction - using existing cache files" -ForegroundColor Green
        Write-Host "[UseExistingCache] Auto-update FORCED to disabled (NoAutoUpdate=TRUE) to preserve patched version" -ForegroundColor Green
        
        # Set minimal platform info (needed for some functions)
        $PlatOS = 'Windows'
        
        # IMPORTANT: If Policy or Advisor are NOT skipped, we still need to authenticate to collect this data
        # Policy and Advisor require API calls and cannot be loaded from resource cache files
        # Handle both switch parameters and boolean values
        $skipPolicyCheck = if ($SkipPolicy -is [switch]) { $SkipPolicy.IsPresent } else { $SkipPolicy -eq $true }
        $skipAdvisoryCheck = if ($SkipAdvisory -is [switch]) { $SkipAdvisory.IsPresent } else { $SkipAdvisory -eq $true }
        $needAuthForPolicyOrAdvisor = (-not $skipPolicyCheck) -or (-not $skipAdvisoryCheck)
        
        if ($needAuthForPolicyOrAdvisor) {
            Write-Host "[UseExistingCache] Policy/Advisor collection requested - authenticating to Azure for API calls" -ForegroundColor Yellow
            # Still need to authenticate for Policy/Advisor collection
            $PlatOS = Test-ARIPS
            
            if ($PlatOS -ne 'Azure CloudShell' -and !$Automation.IsPresent) {
                $TenantID = Connect-ARILoginSession -AzureEnvironment $AzureEnvironment -TenantID $TenantID -SubscriptionID $SubscriptionID -DeviceLogin $DeviceLogin -AppId $AppId -Secret $Secret -CertificatePath $CertificatePath
                
                if (!$NoAutoUpdate.IsPresent) {
                    # Auto-update logic here (but we've forced NoAutoUpdate, so this won't run)
                }
            }
            
            # Get subscriptions for Policy/Advisor collection
            if ([string]::IsNullOrEmpty($SubscriptionID)) {
                $Subscriptions = Get-ARISubscriptions -TenantID $TenantID -AzureEnvironment $AzureEnvironment
            } else {
                $Subscriptions = Get-ARISubscriptions -TenantID $TenantID -SubscriptionID $SubscriptionID -AzureEnvironment $AzureEnvironment
            }
            
            Write-Host "[UseExistingCache] Authenticated and ready to collect Policy/Advisor data" -ForegroundColor Green
        } else {
            # Create empty subscriptions array (some reporting functions may reference it)
            $Subscriptions = @()
            Write-Host "[UseExistingCache] Skipped authentication (Policy and Advisor are skipped)" -ForegroundColor Green
        }
    } else {
        $PlatOS = Test-ARIPS

        if ($PlatOS -ne 'Azure CloudShell' -and !$Automation.IsPresent)
            {
                $TenantID = Connect-ARILoginSession -AzureEnvironment $AzureEnvironment -TenantID $TenantID -SubscriptionID $SubscriptionID -DeviceLogin $DeviceLogin -AppId $AppId -Secret $Secret -CertificatePath $CertificatePath

                if (!$NoAutoUpdate.IsPresent)
                    {
                        Write-Host ('Checking for Powershell Module Updates..')
                        Update-Module -Name AzureResourceInventory -AcceptLicense
                    }
            }
        elseif ($Automation.IsPresent)
            {
                try {
                    $AzureConnection = (Connect-AzAccount -Identity).context

                    Set-AzContext -SubscriptionName $AzureConnection.Subscription -DefaultProfile $AzureConnection
                }
                catch {
                    Write-Output "Failed to set Automation Account requirements. Aborting." 
                    exit
                }
            }

        if ($PlatOS -eq 'Azure CloudShell')
            {
                $Heavy = $true
                $SkipAPIs = $true
            }

        if ($StorageAccount)
            {
                $StorageContext = New-AzStorageContext -StorageAccountName $StorageAccount -UseConnectedAccount
            }

        $Subscriptions = Get-ARISubscriptions -TenantID $TenantID -SubscriptionID $SubscriptionID -PlatOS $PlatOS
    }

    $ReportingPath = Set-ARIReportPath -ReportDir $ReportDir

    $DefaultPath = $ReportingPath.DefaultPath
    $DiagramCache = $ReportingPath.DiagramCache
    $ReportCache = $ReportingPath.ReportCache

    if ($Automation.IsPresent)
        {
            $ReportName = 'ARI_Automation'
        }

    Set-ARIFolder -DefaultPath $DefaultPath -DiagramCache $DiagramCache -ReportCache $ReportCache

    if ($UseExistingCache.IsPresent) {
        Write-Host "[UseExistingCache] Skipping cache clearing and extraction - using existing cache files" -ForegroundColor Green
        
        # Check if cache files exist
        $cacheFiles = @(Get-ChildItem -Path $ReportCache -Filter "*.json" -ErrorAction SilentlyContinue)
        $cacheFileCount = if ($null -ne $cacheFiles) { $cacheFiles.Count } else { 0 }
        if ($cacheFileCount -eq 0) {
            Write-Host "[UseExistingCache] Warning: No cache files found in $ReportCache" -ForegroundColor Yellow
            Write-Host "[UseExistingCache] Excel generation may fail or produce empty report" -ForegroundColor Yellow
        } else {
            Write-Host "[UseExistingCache] Found $cacheFileCount cache file(s) - proceeding to Excel generation" -ForegroundColor Green
        }
        
        # Initialize empty variables for reporting phase (some may be needed)
        $Resources = @()
        $Quotas = $null
        $CostData = $null
        $ResourceContainers = @()
        $Advisories = @()
        $ResourcesCount = 0
        $AdvisoryCount = 0
        $SecCenterCount = 0
        $Security = $null
        $Retirements = @()
        $PolicyCount = 0
        $PolicyAssign = @()
        $PolicyDef = @()
        $PolicySetDef = @()
        
        # Try to load Policy and Advisory data from cache files first
        # Note: PolicyRaw.json is NOT used - Policy data is collected via API call instead
        # This avoids merge issues with PolicyRaw.json files that have inconsistent structures
        $policyCacheFile = Join-Path $ReportCache 'Policy.json'
        $advisoryCacheFile = Join-Path $ReportCache 'Advisory.json'
        $advisoryRawCacheFile = Join-Path $ReportCache 'AdvisoryRaw.json'
        
        # Load Policy data from cache if available
        # Handle both switch parameter and boolean value
        $skipPolicyValue = if ($SkipPolicy -is [switch]) { $SkipPolicy.IsPresent } else { $SkipPolicy -eq $true }
        if (-not $skipPolicyValue) {
            if (Test-Path $policyCacheFile) {
                Write-Host "[UseExistingCache] Loading Policy data from cache file: $policyCacheFile" -ForegroundColor Cyan
                try {
                    $policyCacheData = Get-Content $policyCacheFile -Raw | ConvertFrom-Json
                    
                    # Check if Policy.json contains raw Policy data (PolicyAssign/PolicyDef/PolicySetDef) or processed results
                    if ($policyCacheData.PSObject.Properties['PolicyAssign'] -or $policyCacheData.PSObject.Properties['PolicyDef']) {
                        # Raw Policy data structure - extract it
                        $PolicyAssign = $policyCacheData.PolicyAssign
                        $PolicyDef = $policyCacheData.PolicyDef
                        $PolicySetDef = $policyCacheData.PolicySetDef
                        
                        # Ensure arrays
                        if ($null -eq $PolicyDef) { $PolicyDef = @() }
                        if ($null -eq $PolicySetDef) { $PolicySetDef = @() }
                        
                        # Get count for logging
                        $policyCount = 0
                        if ($null -ne $PolicyAssign) {
                            if ($PolicyAssign -is [PSCustomObject] -or $PolicyAssign -is [System.Collections.Hashtable]) {
                                if ($PolicyAssign.policyAssignments -is [System.Array]) {
                                    $policyCount = $PolicyAssign.policyAssignments.Count
                                }
                            } elseif ($PolicyAssign -is [System.Array]) {
                                $policyCount = $PolicyAssign.Count
                            }
                        }
                        
                        Write-Host "[UseExistingCache] Loaded Policy data from cache ($policyCount assignment(s), $($PolicyDef.Count) definition(s), $($PolicySetDef.Count) set definition(s))" -ForegroundColor Green
                        $hasPolicyData = $true
                    } else {
                        # Processed Policy results (array) - mark as available but will be processed by Policy job
                        $policyCacheCount = 0
                        if ($null -ne $policyCacheData) {
                            if ($policyCacheData -is [System.Array]) {
                                $policyCacheCount = $policyCacheData.Count
                            } elseif ($policyCacheData -is [PSCustomObject]) {
                                $policyCacheCount = 1
                            } else {
                                try {
                                    $policyCacheCount = $policyCacheData.Count
                                } catch {
                                    $policyCacheCount = 1
                                }
                            }
                        }
                        Write-Host "[UseExistingCache] Loaded Policy cache file ($policyCacheCount policy record(s))" -ForegroundColor Green
                        # Policy data will be loaded when Start-ARIExtraJobs processes the Policy job
                    }
                } catch {
                    Write-Host "[UseExistingCache] Warning: Failed to load Policy cache file: $_" -ForegroundColor Yellow
                }
            } else {
                Write-Host "[UseExistingCache] Policy cache file not found - will collect via API call" -ForegroundColor Gray
            }
            
            # PolicyRaw.json loading removed - Policy data collected via API call instead
            # This avoids merge issues with PolicyRaw.json files that have inconsistent structures
            if ($false) {
                Write-Host "[UseExistingCache] Loading raw Policy data from cache file: $policyRawCacheFile" -ForegroundColor Cyan
                try {
                    $policyRawData = Get-Content $policyRawCacheFile -Raw | ConvertFrom-Json
                    $PolicyAssign = $policyRawData.PolicyAssign
                    $PolicyDef = $policyRawData.PolicyDef
                    $PolicySetDef = $policyRawData.PolicySetDef
                    
                    # Handle PolicyAssign structure - it may be an object with policyAssignments property, or a direct array
                    if ($null -ne $PolicyAssign) {
                        # Ensure PolicyAssign has the expected structure
                        if ($PolicyAssign -is [PSCustomObject] -or $PolicyAssign -is [System.Collections.Hashtable]) {
                            # Already has structure, check for policyAssignments property
                            if (-not $PolicyAssign.policyAssignments) {
                                # Convert to hashtable with policyAssignments property
                                $PolicyAssign = @{ policyAssignments = @() }
                            }
                        } elseif ($PolicyAssign -is [System.Array]) {
                            # Direct array - wrap in hashtable with policyAssignments property
                            $PolicyAssign = @{ policyAssignments = $PolicyAssign }
                        } else {
                            # Single value - wrap in hashtable
                            $PolicyAssign = @{ policyAssignments = @($PolicyAssign) }
                        }
                    } else {
                        $PolicyAssign = @{ policyAssignments = @() }
                    }
                    
                    # Safely get count
                    if ($PolicyAssign.policyAssignments -is [System.Array]) {
                        $PolicyCount = [string]$PolicyAssign.policyAssignments.Count
                    } elseif ($null -ne $PolicyAssign.policyAssignments) {
                        $PolicyCount = "1"
                    } else {
                        $PolicyCount = "0"
                    }
                    Write-Host "[UseExistingCache] Loaded raw Policy data ($PolicyCount policy assignment(s))" -ForegroundColor Green
                } catch {
                    Write-Host "[UseExistingCache] Warning: Failed to load raw Policy cache file: $_" -ForegroundColor Yellow
                    $PolicyAssign = @{ policyAssignments = @() }
                    $PolicyDef = @()
                    $PolicySetDef = @()
                }
            }
        }
        
        # Advisory data is NOT loaded from cache - it will be collected via API call instead
        # Handle both switch parameter and boolean value
        $skipAdvisoryValue = if ($SkipAdvisory -is [switch]) { $SkipAdvisory.IsPresent } else { $SkipAdvisory -eq $true }
        if (-not $skipAdvisoryValue) {
            Write-Host "[UseExistingCache] Advisory cache file not found - will collect via API call" -ForegroundColor Gray
        }
        
        # Policy and Advisor data will be collected below if authentication is available AND cache files don't exist
        
        # Extract subscription information from SubscriptionID parameter for Overview and Subscriptions sheets
        # Create minimal subscription objects with Name and Id properties
        # Safely check Count property
        $subscriptionIDCount = if ($null -ne $SubscriptionID -and $SubscriptionID -is [System.Array]) { $SubscriptionID.Count } elseif ($null -ne $SubscriptionID) { 1 } else { 0 }
        if ($subscriptionIDCount -gt 0) {
            Write-Host "[UseExistingCache] Creating subscription objects from SubscriptionID parameter" -ForegroundColor Green
            $Subscriptions = @()
            foreach ($subId in $SubscriptionID) {
                # Create minimal subscription object
                # Since we're using existing cache, we may not have Azure authentication
                # Try to get subscription name from Azure if authenticated, otherwise use ID as name
                $subName = $subId
                try {
                    # Check if we're authenticated to Azure
                    $context = Get-AzContext -ErrorAction SilentlyContinue
                    if ($context) {
                        $azSub = Get-AzSubscription -SubscriptionId $subId -ErrorAction SilentlyContinue
                        if ($azSub -and $azSub.Name) {
                            $subName = $azSub.Name
                        }
                    }
                } catch {
                    # If we can't get subscription info (not authenticated or other error), use ID as name
                    # This is fine - the Subscriptions sheet will still work with subscription IDs
                }
                
                $Subscriptions += [PSCustomObject]@{
                    Id = $subId
                    Name = $subName
                }
            }
            $subscriptionsCreatedCount = if ($null -ne $Subscriptions -and $Subscriptions -is [System.Array]) { $Subscriptions.Count } elseif ($null -ne $Subscriptions) { 1 } else { 0 }
            Write-Host "[UseExistingCache] Created $subscriptionsCreatedCount subscription object(s)" -ForegroundColor Green
        }
        
        # Extract resource data from cache files for Subscriptions sheet
        # The Subscriptions sheet needs resources with: id, Type, location, resourcegroup, subscriptionid
        Write-Host "[UseExistingCache] Extracting resource data from cache files for Subscriptions sheet" -ForegroundColor Green
        $allResources = @()
        foreach ($cacheFile in $cacheFiles) {
            try {
                $cacheContent = Get-Content $cacheFile.FullName -Raw | ConvertFrom-Json
                # Cache files can be either PSCustomObject with properties containing arrays, or direct arrays
                $resourceArrays = @()
                if ($cacheContent -is [PSCustomObject]) {
                    foreach ($prop in $cacheContent.PSObject.Properties) {
                        if ($prop.Value -is [System.Array]) {
                            $resourceArrays += $prop.Value
                        }
                    }
                } elseif ($cacheContent -is [System.Array]) {
                    $resourceArrays = $cacheContent
                }
                
                # Extract resource information from processed cache data
                foreach ($resource in $resourceArrays) {
                    if ($resource -is [PSCustomObject]) {
                        # Try to find resource ID property (could be 'ID', 'Id', 'id', etc.)
                        $resourceId = $null
                        $idProps = @('ID', 'Id', 'id', 'ResourceId', 'resourceId')
                        foreach ($prop in $idProps) {
                            if ($resource.PSObject.Properties.Name -contains $prop) {
                                $resourceId = $resource.$prop
                                break
                            }
                        }
                        
                        if ($resourceId -and $resourceId -match '/subscriptions/([^/]+)/') {
                            $subId = $matches[1]
                            
                            # Try to find Type property (cache files may use 'Resource Type' with space, or 'Type', 'type', etc.)
                            $resourceType = 'Unknown'
                            $typeProps = @('Resource Type', 'Type', 'type', 'ResourceType', 'resourceType', 'TYPE')
                            foreach ($prop in $typeProps) {
                                if ($resource.PSObject.Properties.Name -contains $prop) {
                                    $resourceType = $resource.$prop
                                    break
                                }
                            }
                            
                            # If still unknown, try to extract from resource ID
                            if ($resourceType -eq 'Unknown' -and $resourceId) {
                                if ($resourceId -match '/providers/([^/]+/[^/]+)') {
                                    $resourceType = $matches[1]
                                }
                            }
                            
                            # Try to find Location property
                            $location = ''
                            $locProps = @('Location', 'location', 'LOCATION')
                            foreach ($prop in $locProps) {
                                if ($resource.PSObject.Properties.Name -contains $prop) {
                                    $location = $resource.$prop
                                    break
                                }
                            }
                            
                            # Try to find Resource Group property
                            $resourceGroup = ''
                            $rgProps = @('Resource Group', 'RESOURCEGROUP', 'ResourceGroup', 'resourceGroup', 'resourcegroup')
                            foreach ($prop in $rgProps) {
                                if ($resource.PSObject.Properties.Name -contains $prop) {
                                    $resourceGroup = $resource.$prop
                                    break
                                }
                            }
                            
                            # Create resource object with standard property names
                            # Note: PowerShell hashtables are case-insensitive, so we can't have both 'id' and 'Id'
                            # We'll use standard casing that ARI expects
                            $resourceObj = [PSCustomObject]@{
                                id = $resourceId
                                Type = $resourceType
                                location = $location
                                resourcegroup = $resourceGroup
                                'Resource Group' = $resourceGroup
                                subscriptionid = $subId
                            }
                            $allResources += $resourceObj
                        }
                    }
                }
            } catch {
                Write-Debug "[UseExistingCache] Error reading cache file $($cacheFile.Name): $_"
            }
        }
        $Resources = if ($null -ne $allResources) { $allResources } else { @() }
        $resourceCount = if ($null -ne $Resources -and $Resources -is [System.Array]) { $Resources.Count } elseif ($null -ne $Resources) { 1 } else { 0 }
        Write-Host "[UseExistingCache] Extracted $resourceCount resource(s) from cache files" -ForegroundColor Green
        
        # Collect Policy and Advisor data if not skipped (requires authentication)
        # Safely check Subscriptions.Count
        $subscriptionsCount = if ($null -ne $Subscriptions -and $Subscriptions -is [System.Array]) { $Subscriptions.Count } elseif ($null -ne $Subscriptions) { 1 } else { 0 }
        if ($needAuthForPolicyOrAdvisor -and $subscriptionsCount -gt 0) {
            Write-Host "[UseExistingCache] Collecting Policy and Advisor data via API calls..." -ForegroundColor Yellow
            
            # Collect Advisor data if not skipped AND not already loaded from cache
            # Safely check Advisories.Count
            $advisoriesCount = if ($null -ne $Advisories -and $Advisories -is [System.Array]) { $Advisories.Count } elseif ($null -ne $Advisories) { 1 } else { 0 }
            if (-not $skipAdvisoryValue -and $advisoriesCount -eq 0) {
                Write-Host "[UseExistingCache] Collecting Advisor data via API (cache file not found)..." -ForegroundColor Cyan
                
                # Aggressive memory cleanup BEFORE Advisor API call
                Write-Host "[UseExistingCache] Running extreme memory cleanup before Advisor API call..." -ForegroundColor Gray
                [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $false)
                [System.GC]::WaitForPendingFinalizers()
                [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                Start-Sleep -Milliseconds 500
                
                try {
                    # Create a switch parameter for SkipAdvisory (Start-ARIGraphExtraction expects a switch)
                    $skipAdvisorySwitch = [switch]$false
                    $GraphData = Start-ARIGraphExtraction -ManagementGroup $null -Subscriptions $Subscriptions -SubscriptionID $SubscriptionID -ResourceGroup $null -SecurityCenter $SecurityCenter -SkipAdvisory:$skipAdvisorySwitch -IncludeTags $IncludeTags -TagKey $TagKey -TagValue $TagValue -AzureEnvironment $AzureEnvironment
                    $Advisories = $GraphData.Advisories
                    # Ensure Advisories is an array for safe Count access
                    if ($null -eq $Advisories) {
                        $Advisories = @()
                    } elseif ($Advisories -isnot [System.Array]) {
                        $Advisories = @($Advisories)
                    }
                    # Safely access Count property - handle null/empty cases
                    $AdvisoryCount = if ($null -ne $Advisories -and $Advisories -is [System.Array]) { [string]$Advisories.Count } else { "0" }
                    Write-Host "[UseExistingCache] Collected $AdvisoryCount Advisor recommendation(s) via API call" -ForegroundColor Green
                    Remove-Variable -Name GraphData -ErrorAction SilentlyContinue
                    
                    # Aggressive memory cleanup after Advisor API call
                    Write-Host "[UseExistingCache] Running memory cleanup after Advisor API call..." -ForegroundColor Gray
                    [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $false)
                    [System.GC]::WaitForPendingFinalizers()
                    [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                    Start-Sleep -Milliseconds 500
                } catch {
                    Write-Host "[UseExistingCache] Warning: Failed to collect Advisor data: $_" -ForegroundColor Yellow
                    $Advisories = @()
                    $AdvisoryCount = 0
                }
            } elseif (-not $skipAdvisoryValue -and $null -ne $Advisories) {
                # Ensure Advisories is an array for safe Count access
                if ($Advisories -isnot [System.Array]) {
                    $Advisories = @($Advisories)
                }
                $advisoriesCountCheck = if ($null -ne $Advisories -and $Advisories -is [System.Array]) { $Advisories.Count } else { 0 }
                if ($advisoriesCountCheck -gt 0) {
                    Write-Host "[UseExistingCache] Using Advisory data from cache file" -ForegroundColor Green
                }
            }
            
            # Collect Policy data if not skipped AND not already loaded from cache
            # Safely check PolicyAssign structure
            $hasPolicyData = $false
            if ($null -ne $PolicyAssign) {
                if ($PolicyAssign -is [PSCustomObject] -or $PolicyAssign -is [System.Collections.Hashtable]) {
                    if ($PolicyAssign.policyAssignments -is [System.Array] -and $PolicyAssign.policyAssignments.Count -gt 0) {
                        $hasPolicyData = $true
                    }
                } elseif ($PolicyAssign -is [System.Array] -and $PolicyAssign.Count -gt 0) {
                    $hasPolicyData = $true
                }
            }
            
            if (-not $skipPolicyValue -and -not $hasPolicyData) {
                # Check available memory before attempting Policy collection
                # If we're already low on memory, skip Policy collection to avoid OOM
                $skipPolicyDueToMemory = $false
                try {
                    # Try Linux /proc/meminfo first (Windmill runs on Linux)
                    if (Test-Path "/proc/meminfo") {
                        $memInfo = Get-Content "/proc/meminfo" | Select-String -Pattern "MemAvailable|MemFree"
                        if ($memInfo) {
                            $memAvailableLine = $memInfo | Where-Object { $_ -match "MemAvailable" }
                            if (-not $memAvailableLine) {
                                $memAvailableLine = $memInfo | Where-Object { $_ -match "MemFree" }
                            }
                            if ($memAvailableLine -match "(\d+)") {
                                $availableKB = [int]$matches[1]
                                $availableMB = [math]::Round($availableKB / 1024, 2)
                                Write-Host "[UseExistingCache] Available memory: $availableMB MB" -ForegroundColor Gray
                                
                                # Skip Policy collection if less than 500MB free (conservative threshold for Windmill)
                                if ($availableMB -lt 500) {
                                    Write-Host "[UseExistingCache] WARNING: Low memory detected ($availableMB MB free). Skipping Policy collection to prevent OOM." -ForegroundColor Yellow
                                    Write-Host "[UseExistingCache] Policy data will not be included in this report. Consider collecting Policy data during batch processing." -ForegroundColor Yellow
                                    $skipPolicyDueToMemory = $true
                                }
                            }
                        }
                    } else {
                        # Try WMI (Windows)
                        $memInfo = Get-WmiObject Win32_OperatingSystem -ErrorAction SilentlyContinue
                        if ($null -ne $memInfo) {
                            $availableMB = [math]::Round($memInfo.FreePhysicalMemory / 1024, 2)
                            $totalMB = [math]::Round($memInfo.TotalVisibleMemorySize / 1024, 2)
                            $percentFree = [math]::Round(($availableMB / $totalMB) * 100, 2)
                            Write-Host "[UseExistingCache] Available memory: $availableMB MB ($percentFree% free)" -ForegroundColor Gray
                            
                            if ($availableMB -lt 500) {
                                Write-Host "[UseExistingCache] WARNING: Low memory detected ($availableMB MB free). Skipping Policy collection to prevent OOM." -ForegroundColor Yellow
                                Write-Host "[UseExistingCache] Policy data will not be included in this report. Consider collecting Policy data during batch processing." -ForegroundColor Yellow
                                $skipPolicyDueToMemory = $true
                            }
                        }
                    }
                } catch {
                    Write-Host "[UseExistingCache] Warning: Could not check available memory: $_" -ForegroundColor Yellow
                }
                
                if ($skipPolicyDueToMemory) {
                    $PolicyAssign = @{ policyAssignments = @() }
                    $PolicyDef = @()
                    $PolicySetDef = @()
                    $PolicyCount = 0
                } else {
                    Write-Host "[UseExistingCache] Collecting Policy data via API (cache file not found)..." -ForegroundColor Cyan
                    
                    # EXTREME memory cleanup BEFORE Policy API call (multiple iterations)
                    Write-Host "[UseExistingCache] Running EXTREME memory cleanup before Policy API call..." -ForegroundColor Gray
                    try {
                        Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
                        for ($i = 1; $i -le 10; $i++) {
                            [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $false)
                            [System.GC]::WaitForPendingFinalizers()
                            [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                        }
                        Clear-ARIMemory
                        Start-Sleep -Milliseconds 1000  # Give system time to free memory
                    } catch {
                        Write-Host "[UseExistingCache] Warning: Pre-Policy cleanup had issues: $_" -ForegroundColor Yellow
                    }
                    
                    try {
                        # Create a switch parameter for SkipPolicy (Get-ARIAPIResources expects a switch)
                        $skipPolicySwitch = [switch]$false
                        $APIResults = Get-ARIAPIResources -Subscriptions $Subscriptions -AzureEnvironment $AzureEnvironment -SkipPolicy:$skipPolicySwitch
                        
                        # Collect Resource Health events (needed for Outages sheet)
                        # Resource Health events are collected via Get-ARIAPIResources but not added to Resources when using cache
                        # Get-ARIAPIResources returns an array of hashtables, each with ResourceHealth.value
                        $resourceHealthCount = 0
                        if ($null -ne $APIResults) {
                            if ($APIResults -isnot [System.Array]) {
                                $APIResults = @($APIResults)
                            }
                            foreach ($apiResult in $APIResults) {
                                if ($null -ne $apiResult -and $null -ne $apiResult.ResourceHealth) {
                                    $resourceHealthEvents = $apiResult.ResourceHealth
                                    if ($null -ne $resourceHealthEvents) {
                                        # ResourceHealth is typically an array from API response
                                        if ($resourceHealthEvents -isnot [System.Array]) {
                                            $resourceHealthEvents = @($resourceHealthEvents)
                                        }
                                        # Add Resource Health events to Resources array so Outages module can process them
                                        foreach ($event in $resourceHealthEvents) {
                                            if ($null -ne $event) {
                                                # Ensure event has TYPE property set correctly for Outages module filtering
                                                # Outages.ps1 filters: TYPE -eq 'Microsoft.ResourceHealth/events'
                                                if ($event -is [PSCustomObject]) {
                                                    # Add TYPE property if it doesn't exist
                                                    if (-not ($event.PSObject.Properties.Name -contains 'TYPE')) {
                                                        $event | Add-Member -MemberType NoteProperty -Name 'TYPE' -Value 'Microsoft.ResourceHealth/events' -Force
                                                    }
                                                } elseif ($event -is [hashtable]) {
                                                    # Convert hashtable to PSCustomObject with TYPE
                                                    $event = [PSCustomObject]@{
                                                        TYPE = 'Microsoft.ResourceHealth/events'
                                                        properties = $event.properties
                                                        name = $event.name
                                                        id = $event.id
                                                    }
                                                } else {
                                                    # Wrap in PSCustomObject
                                                    $event = [PSCustomObject]@{
                                                        TYPE = 'Microsoft.ResourceHealth/events'
                                                        properties = $event
                                                    }
                                                }
                                                $Resources += $event
                                                $resourceHealthCount++
                                            }
                                        }
                                    }
                                }
                            }
                            if ($resourceHealthCount -gt 0) {
                                Write-Host "[UseExistingCache] Added $resourceHealthCount Resource Health event(s) to Resources for Outages processing" -ForegroundColor Green
                                
                                # Store Resource Health events in script-scoped variable for later Outages sheet generation
                                # Filter for outage events (same filter as Outages.ps1)
                                $script:ResourceHealthEventsForOutages = $Resources | Where-Object { 
                                    $_.TYPE -eq 'Microsoft.ResourceHealth/events' -and 
                                    $null -ne $_.properties -and
                                    $null -ne $_.properties.description -and
                                    $_.properties.description -like '*How can customers make incidents like this less impactful?*' 
                                }
                                Write-Host "[UseExistingCache] Stored $($script:ResourceHealthEventsForOutages.Count) outage event(s) for direct sheet generation" -ForegroundColor Green
                            }
                        }
                        
                        # Extract Policy data from APIResults
                        # Get-ARIAPIResources returns array of hashtables
                        # In normal flow, Start-ARIExtractionOrchestration does: $Resources += $APIResults.ResourceHealth
                        # This works because PowerShell automatically expands array properties
                        # But we need to handle it explicitly here
                        $allPolicyAssign = @()
                        $allPolicyDef = @()
                        $allPolicySetDef = @()
                        if ($null -ne $APIResults) {
                            if ($APIResults -isnot [System.Array]) {
                                $APIResults = @($APIResults)
                            }
                            foreach ($apiResult in $APIResults) {
                                if ($null -ne $apiResult.PolicyAssign) {
                                    if ($apiResult.PolicyAssign -is [System.Array]) {
                                        $allPolicyAssign += $apiResult.PolicyAssign
                                    } else {
                                        $allPolicyAssign += $apiResult.PolicyAssign
                                    }
                                }
                                if ($null -ne $apiResult.PolicyDef) {
                                    if ($apiResult.PolicyDef -is [System.Array]) {
                                        $allPolicyDef += $apiResult.PolicyDef
                                    } else {
                                        $allPolicyDef += $apiResult.PolicyDef
                                    }
                                }
                                if ($null -ne $apiResult.PolicySetDef) {
                                    if ($apiResult.PolicySetDef -is [System.Array]) {
                                        $allPolicySetDef += $apiResult.PolicySetDef
                                    } else {
                                        $allPolicySetDef += $apiResult.PolicySetDef
                                    }
                                }
                            }
                        }
                        # Set Policy variables (keep existing structure for compatibility)
                        $PolicyAssign = $allPolicyAssign
                        $PolicyDef = $allPolicyDef
                        $PolicySetDef = $allPolicySetDef
                        # Safely access Count property - handle null/empty cases
                        if ($null -ne $PolicyAssign) {
                            if ($PolicyAssign -is [PSCustomObject] -or $PolicyAssign -is [System.Collections.Hashtable]) {
                                if ($PolicyAssign.policyAssignments -is [System.Array]) {
                                    $PolicyCount = [string]$PolicyAssign.policyAssignments.Count
                                } else {
                                    $PolicyCount = "0"
                                }
                            } elseif ($PolicyAssign -is [System.Array]) {
                                $PolicyCount = if ($null -ne $PolicyAssign -and $PolicyAssign -is [System.Array]) { [string]$PolicyAssign.Count } else { "0" }
                            } else {
                                $PolicyCount = "1"
                            }
                        } else {
                            $PolicyCount = "0"
                        }
                        Write-Host "[UseExistingCache] Collected $PolicyCount Policy assignment(s) via API call" -ForegroundColor Green
                        Remove-Variable -Name APIResults -ErrorAction SilentlyContinue
                        
                        # Aggressive memory cleanup after Policy API call
                        Write-Host "[UseExistingCache] Running memory cleanup after Policy API call..." -ForegroundColor Gray
                        [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $false)
                        [System.GC]::WaitForPendingFinalizers()
                        [System.GC]::Collect([System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                        Start-Sleep -Milliseconds 500
                    } catch {
                        Write-Host "[UseExistingCache] Warning: Failed to collect Policy data: $_" -ForegroundColor Yellow
                        $PolicyAssign = @{ policyAssignments = @() }
                        $PolicyDef = @()
                        $PolicySetDef = @()
                        $PolicyCount = 0
                    }
                }
            } elseif (-not $skipPolicyValue -and $hasPolicyData) {
                Write-Host "[UseExistingCache] Using Policy data from cache file" -ForegroundColor Green
            }
        }
        
        # Create a dummy stopwatch for timing (reporting functions may reference it)
        $ExtractionRuntime = [System.Diagnostics.Stopwatch]::StartNew()
        $ExtractionRuntime.Stop()
        $ExtractionTotalTime = "00:00:00:000"
        
        Write-Host "[UseExistingCache] Completed cache loading and Policy/Advisor collection" -ForegroundColor Green
    } else {
        Clear-ARICacheFolder -ReportCache $ReportCache

        Get-Job | Where-Object {$_.name -like 'ResourceJob_*'} | Remove-Job -Force | Out-Null

        $ExtractionRuntime = [System.Diagnostics.Stopwatch]::StartNew()

            $ExtractionData = Start-ARIExtractionOrchestration -ManagementGroup $ManagementGroup -Subscriptions $Subscriptions -SubscriptionID $SubscriptionID -ResourceGroup $ResourceGroup -SecurityCenter $SecurityCenter -SkipAdvisory $SkipAdvisory -SkipPolicy $SkipPolicy -IncludeTags $IncludeTags -TagKey $TagKey -TagValue $TagValue -SkipAPIs $SkipAPIs -SkipVMDetails $SkipVMDetails -IncludeCosts $IncludeCosts -Automation $Automation -AzureEnvironment $AzureEnvironment

        $ExtractionRuntime.Stop()

        $Resources = $ExtractionData.Resources
        $Quotas = $ExtractionData.Quotas
        $CostData = $ExtractionData.Costs
        $ResourceContainers = $ExtractionData.ResourceContainers
        $Advisories = $ExtractionData.Advisories
        $ResourcesCount = $ExtractionData.ResourcesCount
        $AdvisoryCount = $ExtractionData.AdvisoryCount
        $SecCenterCount = $ExtractionData.SecCenterCount
        $Security = $ExtractionData.Security
        $Retirements = $ExtractionData.Retirements
        $PolicyCount = $ExtractionData.PolicyCount
        $PolicyAssign = $ExtractionData.PolicyAssign
        $PolicyDef = $ExtractionData.PolicyDef
        $PolicySetDef = $ExtractionData.PolicySetDef

        Remove-Variable -Name ExtractionData -ErrorAction SilentlyContinue

        $ExtractionTotalTime = $ExtractionRuntime.Elapsed.ToString("dd\:hh\:mm\:ss\:fff")

        if ($Automation.IsPresent)
            {
                Write-Output "Extraction Phase Finished"
                Write-Output ('Total Extraction Time: ' + $ExtractionTotalTime)
            }
        else
            {
                Write-Host "Extraction Phase Finished: " -ForegroundColor Green -NoNewline
                Write-Host $ExtractionTotalTime -ForegroundColor Cyan
            }
    }

    #### Creating Excel file variable:
    $FileName = ($ReportName + "_Report_" + (get-date -Format "yyyy-MM-dd_HH_mm") + ".xlsx")
    $File = Join-Path $DefaultPath $FileName
    #$DFile = ($DefaultPath + $Global:ReportName + "_Diagram_" + (get-date -Format "yyyy-MM-dd_HH_mm") + ".vsdx")
    $DDName = ($ReportName + "_Diagram_" + (get-date -Format "yyyy-MM-dd_HH_mm") + ".xml")
    $DDFile = Join-Path $DefaultPath $DDName 

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Excel file: ' + $File)

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Default Jobs.')

    $ProcessingRunTime = [System.Diagnostics.Stopwatch]::StartNew()

    # Skip processing phase when using existing cache (cache files already exist)
    # But still run Start-ARIExtraJobs to create necessary jobs for reporting (like Subscriptions)
    if ($UseExistingCache.IsPresent) {
        Write-Host "[UseExistingCache] Skipping resource processing - using existing cache files directly" -ForegroundColor Green
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'[UseExistingCache] Skipped Start-ARIProcessOrchestration')
        
        # Ensure Resources is initialized and is an array for the subscription job
        if ($null -eq $Resources) {
            $Resources = @()
        } elseif ($Resources -isnot [System.Array]) {
            $Resources = @($Resources)
        }
        # Safely get Resources count
        $resourcesCount = if ($null -ne $Resources -and ($Resources -is [System.Array] -or $Resources -is [System.Collections.ICollection])) {
            $Resources.Count
        } elseif ($null -ne $Resources) {
            1
        } else {
            0
        }
        $resourcesForJob = if ($resourcesCount -gt 0) { $Resources } else { @() }
        Write-Debug "[UseExistingCache] Passing $resourcesCount resource(s) to Start-ARIExtraJobs for subscription job"
        
        # Ensure Subscriptions is initialized and is an array for the subscription job
        if ($null -eq $Subscriptions) {
            $Subscriptions = @()
        } elseif ($Subscriptions -isnot [System.Array]) {
            $Subscriptions = @($Subscriptions)
        }
        # Safely get Subscriptions count
        $subscriptionsCountForJob = if ($null -ne $Subscriptions -and ($Subscriptions -is [System.Array] -or $Subscriptions -is [System.Collections.ICollection])) {
            $Subscriptions.Count
        } elseif ($null -ne $Subscriptions) {
            1
        } else {
            0
        }
        Write-Debug "[UseExistingCache] Passing $subscriptionsCountForJob subscription(s) to Start-ARIExtraJobs"
        
        # Still run Start-ARIExtraJobs to create jobs needed for reporting (Subscriptions, etc.)
        # but skip diagram and other resource-intensive jobs
        try {
            Write-Debug "[UseExistingCache] Calling Start-ARIExtraJobs with Subscriptions=$subscriptionsCountForJob, Resources=$resourcesCount"
            Start-ARIExtraJobs -SkipDiagram $SkipDiagram -SkipAdvisory $SkipAdvisory -SkipPolicy $SkipPolicy -SecurityCenter $Security -Subscriptions $Subscriptions -Resources $resourcesForJob -Advisories $Advisories -DDFile $DDFile -DiagramCache $DiagramCache -FullEnv $FullEnv -ResourceContainers $ResourceContainers -Security $Security -PolicyAssign $PolicyAssign -PolicySetDef $PolicySetDef -PolicyDef $PolicyDef -IncludeCosts $IncludeCosts -CostData $CostData -Automation $Automation
            Write-Debug "[UseExistingCache] Start-ARIExtraJobs completed successfully"
        } catch {
            # Safe error handling - check property existence before accessing
            $errorMsg = if ($null -ne $_ -and $null -ne $_.Exception) { $_.Exception.Message } else { "Unknown error" }
            $errorStack = if ($null -ne $_ -and $null -ne $_.ScriptStackTrace) { $_.ScriptStackTrace } else { "No stack trace available" }
            
            Write-Error "Error in Start-ARIExtraJobs: $errorMsg"
            Write-Error "Stack trace: $errorStack"
            
            $errorLine = "Unknown"
            $errorFunc = "Unknown"
            
            try {
                if ($null -ne $_ -and $null -ne $_.InvocationInfo) {
                    $errorLine = if ($null -ne $_.InvocationInfo.ScriptLineNumber) { $_.InvocationInfo.ScriptLineNumber } else { "Unknown" }
                    # Check if FunctionName property exists before accessing
                    if ($_.InvocationInfo.PSObject.Properties.Name -contains 'FunctionName') {
                        $errorFunc = $_.InvocationInfo.FunctionName
                    }
                }
            } catch {
                # Ignore errors accessing InvocationInfo
            }
            
            Write-Error "Line: $errorLine"
            Write-Error "Function: $errorFunc"
            throw
        }
    } else {
        # Ensure Subscriptions is initialized and is an array (safety check)
        if ($null -eq $Subscriptions) {
            $Subscriptions = @()
        } elseif ($Subscriptions -isnot [System.Array]) {
            $Subscriptions = @($Subscriptions)
        }
        
        Start-ARIExtraJobs -SkipDiagram $SkipDiagram -SkipAdvisory $SkipAdvisory -SkipPolicy $SkipPolicy -SecurityCenter $Security -Subscriptions $Subscriptions -Resources $Resources -Advisories $Advisories -DDFile $DDFile -DiagramCache $DiagramCache -FullEnv $FullEnv -ResourceContainers $ResourceContainers -Security $Security -PolicyAssign $PolicyAssign -PolicySetDef $PolicySetDef -PolicyDef $PolicyDef -IncludeCosts $IncludeCosts -CostData $CostData -Automation $Automation

        Start-ARIProcessOrchestration -Subscriptions $Subscriptions -Resources $Resources -Retirements $Retirements -DefaultPath $DefaultPath -Heavy $Heavy -File $File -InTag $InTag -Automation $Automation
    }

    $ProcessingRunTime.Stop()

    $ProcessingTotalTime = $ProcessingRunTime.Elapsed.ToString("dd\:hh\:mm\:ss\:fff")

    if ($Automation.IsPresent)
        {
            Write-Output "Processing Phase Finished"
            Write-Output ('Total Processing Time: ' + $ProcessingTotalTime)
        }
    else
        {
            Write-Host "Processing Phase Finished: " -ForegroundColor Green -NoNewline
            Write-Host $ProcessingTotalTime -ForegroundColor Cyan
        }

    # Skip Excel generation if SkipExcel is specified (useful for batch collection where only cache is needed)
    if ($SkipExcel.IsPresent) {
        Write-Host "[SkipExcel] Skipping Excel generation - only creating cache files" -ForegroundColor Green
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'[SkipExcel] Skipped Excel report generation')
        
        # Note: Policy and Advisory data are NOT cached - they will be collected via API call during Excel generation
        # This avoids merge issues with PolicyRaw.json and AdvisoryRaw.json files that have inconsistent structures
        # Handle both switch parameters and boolean values
        $skipPolicyCheck = if ($SkipPolicy -is [switch]) { $SkipPolicy.IsPresent } else { $SkipPolicy -eq $true }
        if (-not $skipPolicyCheck) {
            # Policy data will be collected via API call during Excel generation
            # No need to cache it during batch collection
            Write-Host "[SkipExcel] Policy data will be collected via API call during Excel generation (not caching)" -ForegroundColor Gray
        }
        
        $skipAdvisoryCheck = if ($SkipAdvisory -is [switch]) { $SkipAdvisory.IsPresent } else { $SkipAdvisory -eq $true }
        if (-not $skipAdvisoryCheck) {
            # Advisory data will be collected via API call during Excel generation
            # No need to cache it during batch collection
            Write-Host "[SkipExcel] Advisory data will be collected via API call during Excel generation (not caching)" -ForegroundColor Gray
        }
        
        $TotalRes = 0
        $ReportingRunTime = [System.Diagnostics.Stopwatch]::StartNew()
        $ReportingRunTime.Stop()
        $ReportingTotalTime = "00:00:00:00:000"
        if ($Automation.IsPresent) {
            Write-Output "Cache Collection Finished"
            Write-Output ('Total Processing Time: ' + $ReportingTotalTime)
        } else {
            Write-Host "Cache Collection Finished: " -ForegroundColor Green -NoNewline
            Write-Host $ReportingTotalTime -ForegroundColor Cyan
        }
    } else {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Resources Report Function.')
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Excel Table Style used: ' + $TableStyle)

        $ReportingRunTime = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            Start-ARIReporOrchestration -ReportCache $ReportCache -SecurityCenter $SecurityCenter -File $File -Quotas $Quotas -SkipPolicy $SkipPolicy -SkipAdvisory $SkipAdvisory -IncludeCosts $IncludeCosts -Automation $Automation -TableStyle $TableStyle -Advisories $Advisories
        } catch {
            # Safe error handling - check property existence before accessing
            $errorDetails = if ($null -ne $_ -and $null -ne $_.Exception) { $_.Exception.Message } else { "Unknown error" }
            
            $errorLine = "Unknown"
            $errorFunction = "Unknown"
            
            try {
                if ($null -ne $_ -and $null -ne $_.InvocationInfo) {
                    $errorLine = if ($null -ne $_.InvocationInfo.ScriptLineNumber) { $_.InvocationInfo.ScriptLineNumber } else { "Unknown" }
                    # Check if FunctionName property exists before accessing
                    if ($_.InvocationInfo.PSObject.Properties.Name -contains 'FunctionName') {
                        $errorFunction = $_.InvocationInfo.FunctionName
                    }
                }
            } catch {
                # Ignore errors accessing InvocationInfo
            }
            
            $errorStack = if ($null -ne $_ -and $null -ne $_.ScriptStackTrace) { $_.ScriptStackTrace } else { "No stack trace available" }
            
            Write-Error "Excel generation failed in $errorFunction at line $errorLine : $errorDetails"
            Write-Error "Stack trace: $errorStack"
            throw
        }

        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Generating Overview sheet (Charts).')

        # Ensure Subscriptions is initialized and is an array before passing to Start-ARIExcelCustomization
        if ($null -eq $Subscriptions) {
            $Subscriptions = @()
        } elseif ($Subscriptions -isnot [System.Array]) {
            $Subscriptions = @($Subscriptions)
        }

            $TotalRes = Start-ARIExcelCustomization -File $File -TableStyle $TableStyle -PlatOS $PlatOS -Subscriptions $Subscriptions -ExtractionRunTime $ExtractionRuntime -ProcessingRunTime $ProcessingRunTime -ReportingRunTime $ReportingRunTime -IncludeCosts $IncludeCosts -RunLite $RunLite -Overview $Overview

            Write-Progress -activity 'Azure Inventory' -Status "95% Complete." -PercentComplete 95 -CurrentOperation "Excel Customization Completed.."
            
            # Generate Outages sheet directly using working logic from outages-only script
            # This bypasses Outages.ps1 module which has issues with cache data format
            if ($null -ne $script:ResourceHealthEventsForOutages -and $script:ResourceHealthEventsForOutages.Count -gt 0) {
                Write-Host "[UseExistingCache] Generating Outages sheet directly using working logic..." -ForegroundColor Cyan
                try {
                    # Ensure Subscriptions is available and is an array
                    if ($null -eq $Subscriptions) {
                        $Subscriptions = @()
                    } elseif ($Subscriptions -isnot [System.Array]) {
                        $Subscriptions = @($Subscriptions)
                    }
                    
                    # Process outages using same logic as outages-only script
                    $processedOutages = @()
                    $SubObjects = @()
                    foreach ($sub in $Subscriptions) {
                        $SubObjects += [PSCustomObject]@{
                            id = $sub.Id
                            name = $sub.Name
                        }
                    }
                    
                    foreach ($outage in $script:ResourceHealthEventsForOutages) {
                        try {
                            # Safely extract impacted subscriptions
                            $ImpactedSubs = @()
                            if ($null -ne $outage.properties -and $null -ne $outage.properties.impact -and $null -ne $outage.properties.impact.impactedRegions) {
                                if ($null -ne $outage.properties.impact.impactedRegions.impactedSubscriptions) {
                                    $uniqueSubs = $outage.properties.impact.impactedRegions.impactedSubscriptions | Select-Object -Unique
                                    if ($null -ne $uniqueSubs) {
                                        if ($uniqueSubs -is [System.Array]) {
                                            $ImpactedSubs = $uniqueSubs
                                        } else {
                                            $ImpactedSubs = @($uniqueSubs)
                                        }
                                    }
                                }
                            }
                            # If no impacted subscriptions found, try to extract from outage ID
                            $impactedSubsCount = if ($null -ne $ImpactedSubs -and $ImpactedSubs -is [System.Array]) { $ImpactedSubs.Count } elseif ($null -ne $ImpactedSubs) { 1 } else { 0 }
                            if ($impactedSubsCount -eq 0) {
                                if ($null -ne $outage.id -and $outage.id -match '/subscriptions/([^/]+)') {
                                    $ImpactedSubs = @($matches[1])
                                }
                            }
                            
                            $ResUCount = 1
                            $Data = $outage.properties
                            
                            foreach ($Sub0 in $ImpactedSubs) {
                                $sub1 = $SubObjects | Where-Object { $_.id -eq $Sub0 }
                                if ($sub1) {
                                    $StartTime = $Data.impactStartTime
                                    $StartTime = [datetime]$StartTime
                                    $StartTime = $StartTime.ToString("yyyy-MM-dd HH:mm")
                                    
                                    $Mitigation = $Data.impactMitigationTime
                                    $Mitigation = [datetime]$Mitigation
                                    $Mitigation = $Mitigation.ToString("yyyy-MM-dd HH:mm")
                                    
                                    # Safely handle impactedService
                                    $impactedServiceValue = $outage.properties.impact.impactedService
                                    if ($null -ne $impactedServiceValue) {
                                        if ($impactedServiceValue -is [System.Array] -and $impactedServiceValue.Count -gt 1) {
                                            $ImpactedService = $impactedServiceValue | ForEach-Object { $_ + ' ,' }
                                        } else {
                                            $ImpactedService = $impactedServiceValue
                                        }
                                    } else {
                                        $ImpactedService = ''
                                    }
                                    $ImpactedService = [string]$ImpactedService
                                    $ImpactedService = if ($ImpactedService -like '* ,*') { $ImpactedService -replace ".$" } else { $ImpactedService }
                                    
                                    # Safely parse HTML description
                                    $OutageDescription = ''
                                    $SplitDescription = @('', '', '', '', '', '', '')
                                    try {
                                        $HTML = New-Object -Com 'HTMLFile'
                                        $HTML.write([ref]$outage.properties.description)
                                        $OutageDescription = $Html.body.innerText
                                        $SplitDescription = $OutageDescription.split('How can we make our incident communications more useful?').split('How can customers make incidents like this less impactful?').split('How are we making incidents like this less likely or less impactful?').split('How did we respond?').split('What went wrong and why?').split('What happened?')
                                    } catch {
                                        $OutageDescription = $outage.properties.description
                                        $SplitDescription = @('', $OutageDescription, '', '', '', '', '')
                                    }
                                    
                                    # Safely extract split description sections
                                    $whatHappened = ''
                                    $whatWentWrong = ''
                                    $howDidWeRespond = ''
                                    $howMakingLessLikely = ''
                                    $howCustomersCanMakeLessImpactful = ''
                                    
                                    $splitDescCount = if ($null -ne $SplitDescription -and $SplitDescription -is [System.Array]) { $SplitDescription.Count } elseif ($null -ne $SplitDescription) { 1 } else { 0 }
                                    
                                    if ($splitDescCount -gt 1 -and $null -ne $SplitDescription[1]) {
                                        $whatHappenedLines = $SplitDescription[1].Split([Environment]::NewLine)
                                        $whatHappenedLinesCount = if ($null -ne $whatHappenedLines -and $whatHappenedLines -is [System.Array]) { $whatHappenedLines.Count } elseif ($null -ne $whatHappenedLines) { 1 } else { 0 }
                                        if ($whatHappenedLinesCount -gt 1) { $whatHappened = $whatHappenedLines[1] }
                                    }
                                    if ($splitDescCount -gt 2 -and $null -ne $SplitDescription[2]) {
                                        $whatWentWrongLines = $SplitDescription[2].Split([Environment]::NewLine)
                                        $whatWentWrongLinesCount = if ($null -ne $whatWentWrongLines -and $whatWentWrongLines -is [System.Array]) { $whatWentWrongLines.Count } elseif ($null -ne $whatWentWrongLines) { 1 } else { 0 }
                                        if ($whatWentWrongLinesCount -gt 1) { $whatWentWrong = $whatWentWrongLines[1] }
                                    }
                                    if ($splitDescCount -gt 3 -and $null -ne $SplitDescription[3]) {
                                        $howDidWeRespondLines = $SplitDescription[3].Split([Environment]::NewLine)
                                        $howDidWeRespondLinesCount = if ($null -ne $howDidWeRespondLines -and $howDidWeRespondLines -is [System.Array]) { $howDidWeRespondLines.Count } elseif ($null -ne $howDidWeRespondLines) { 1 } else { 0 }
                                        if ($howDidWeRespondLinesCount -gt 1) { $howDidWeRespond = $howDidWeRespondLines[1] }
                                    }
                                    if ($splitDescCount -gt 4 -and $null -ne $SplitDescription[4]) {
                                        $howMakingLessLikelyLines = $SplitDescription[4].Split([Environment]::NewLine)
                                        $howMakingLessLikelyLinesCount = if ($null -ne $howMakingLessLikelyLines -and $howMakingLessLikelyLines -is [System.Array]) { $howMakingLessLikelyLines.Count } elseif ($null -ne $howMakingLessLikelyLines) { 1 } else { 0 }
                                        if ($howMakingLessLikelyLinesCount -gt 1) { $howMakingLessLikely = $howMakingLessLikelyLines[1] }
                                    }
                                    if ($splitDescCount -gt 5 -and $null -ne $SplitDescription[5]) {
                                        $howCustomersCanMakeLessImpactfulLines = $SplitDescription[5].Split([Environment]::NewLine)
                                        $howCustomersCanMakeLessImpactfulLinesCount = if ($null -ne $howCustomersCanMakeLessImpactfulLines -and $howCustomersCanMakeLessImpactfulLines -is [System.Array]) { $howCustomersCanMakeLessImpactfulLines.Count } elseif ($null -ne $howCustomersCanMakeLessImpactfulLines) { 1 } else { 0 }
                                        if ($howCustomersCanMakeLessImpactfulLinesCount -gt 1) { $howCustomersCanMakeLessImpactful = $howCustomersCanMakeLessImpactfulLines[1] }
                                    }
                                    
                                    $obj = [PSCustomObject]@{
                                        'Subscription' = $sub1.name
                                        'Outage ID' = $outage.name
                                        'Event Type' = $Data.eventType
                                        'Status' = $Data.status
                                        'Event Level' = $Data.eventlevel
                                        'Title' = $Data.title
                                        'Impact Start Time' = $StartTime
                                        'Impact Mitigation Time' = $Mitigation
                                        'Impacted Services' = $ImpactedService
                                        'What happened' = $whatHappened
                                        'What went wrong and why' = $whatWentWrong
                                        'How did we respond' = $howDidWeRespond
                                        'How are we making incidents like this less likely or less impactful' = $howMakingLessLikely
                                        'How can customers make incidents like this less impactful' = $howCustomersCanMakeLessImpactful
                                        'Resource U' = $ResUCount
                                    }
                                    $processedOutages += $obj
                                    if ($ResUCount -eq 1) { $ResUCount = 0 }
                                }
                            }
                        } catch {
                            Write-Debug "[UseExistingCache] Error processing outage $($outage.name): $_"
                        }
                    }
                    
                    if ($processedOutages.Count -gt 0) {
                        # Generate Outages sheet
                        $ResourceUSum = ($processedOutages | Measure-Object -Property 'Resource U' -Sum).Sum
                        $TableName = ('OutageTab_'+$ResourceUSum)
                        
                        $Style = @(
                            New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0' -Range 'A:E'
                            New-ExcelStyle -HorizontalAlignment Left -NumberFormat '0' -WrapText -Width 55 -Range 'F:F'
                            New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0' -Range 'G:I'
                            New-ExcelStyle -HorizontalAlignment Left -NumberFormat '0' -WrapText -Width 80 -Range 'J:N'
                        )
                        
                        $Exc = New-Object System.Collections.Generic.List[System.Object]
                        $Exc.Add('Subscription')
                        $Exc.Add('Outage ID')
                        $Exc.Add('Event Type')
                        $Exc.Add('Status')
                        $Exc.Add('Event Level')
                        $Exc.Add('Title')
                        $Exc.Add('Impact Start Time')
                        $Exc.Add('Impact Mitigation Time')
                        $Exc.Add('Impacted Services')
                        $Exc.Add('What happened')
                        $Exc.Add('What went wrong and why')
                        $Exc.Add('How did we respond')
                        $Exc.Add('How are we making incidents like this less likely or less impactful')
                        $Exc.Add('How can customers make incidents like this less impactful')
                        
                        $processedOutages | Select-Object $Exc | Export-Excel -Path $File -WorksheetName 'Outages' -AutoSize -TableName $TableName -MaxAutoSizeRows 100 -TableStyle $TableStyle -Numberformat '0' -Style $Style
                        
                        Write-Host "[UseExistingCache] ✓ Generated Outages sheet with $($processedOutages.Count) outage record(s)" -ForegroundColor Green
                    } else {
                        Write-Host "[UseExistingCache] ⚠ No processed outages to generate sheet" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "[UseExistingCache] ⚠ Error generating Outages sheet: $_" -ForegroundColor Yellow
                    Write-Host "[UseExistingCache] Error details: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

        $ReportingRunTime.Stop()

        $ReportingTotalTime = $ReportingRunTime.Elapsed.ToString("dd\:hh\:mm\:ss\:fff")

        if ($Automation.IsPresent)
            {
                Write-Output "Report Building Finished"
                Write-Output ('Total Processing Time: ' + $ReportingTotalTime)
            }
        else
            {
                Write-Host "Report Building Finished: " -ForegroundColor Green -NoNewline
                Write-Host $ReportingTotalTime -ForegroundColor Cyan
            }
    }

    # Clear memory to remove as many memory footprint as possible
    Clear-ARIMemory

    # Clear Cache Folder for future runs (skip if using existing cache)
    if (-not $UseExistingCache.IsPresent) {
        Clear-ARICacheFolder -ReportCache $ReportCache
    } else {
        Write-Host "[UseExistingCache] Preserving cache files for future use" -ForegroundColor Green
    }

    # Kills any automated Excel process that might be running
    Remove-ARIExcelProcess

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Finished Charts Phase.')

    if(!$SkipDiagram.IsPresent -and !$Automation.IsPresent)
    {
        Write-Progress -activity 'Diagrams' -Status "Completing Diagram" -PercentComplete 70 -CurrentOperation "Consolidating Diagram"

        $JobNames = (Get-Job | Where-Object {$_.name -eq 'DrawDiagram'}).Name

        Wait-ARIJob -JobNames $JobNames -JobType 'Diagram' -LoopTime 5

        Remove-Job -Name 'DrawDiagram' | Out-Null

        Write-Progress -activity 'Diagrams' -Status "Closing Diagram File" -Completed
    }


    if ($StorageAccount)
        {
            Write-Output "Sending Excel file to Storage Account:"
            Write-Output $File
            Set-AzStorageBlobContent -File $File -Container $StorageContainer -Context $StorageContext | Out-Null
            if(!$SkipDiagram.IsPresent)
                {
                    Write-Output "Sending Diagram file to Storage Account:"
                    Write-Output $DDFile
                    Set-AzStorageBlobContent -File $DDFile -Container $StorageContainer -Context $StorageContext | Out-Null
                    if($Debug.IsPresent)
                        {
                            $LogFilePath = Join-Path $DefaultPath 'DiagramLogFile.log'
                            Set-AzStorageBlobContent -File $LogFilePath -Container $StorageContainer -Context $StorageContext -Force | Out-Null
                        }
                }
        }

    $TotalRunTime.Stop()

    $Measure = $TotalRunTime.Elapsed.ToString("dd\:hh\:mm\:ss\:fff")

    Write-Progress -activity 'Azure Inventory' -Status "100% Complete." -Completed

    if (-not $SkipExcel.IsPresent) {
        Out-ARIReportResults -Measure $Measure -ResourcesCount $ResourcesCount -TotalRes $TotalRes -SkipAdvisory $SkipAdvisory -AdvisoryData $AdvisoryCount -SkipPolicy $SkipPolicy -SkipAPIs $SkipAPIs -PolicyData $PolicyCount -SecurityCenter $SecurityCenter -SecurityCenterData $SecCenterCount -File $File -SkipDiagram $SkipDiagram -DDFile $DDFile
    } else {
        Write-Host "Cache collection complete. Excel generation skipped." -ForegroundColor Green
        Write-Host "Total Resources on Azure: $ResourcesCount" -ForegroundColor Cyan
        Write-Host "Cache files saved to: $ReportCache" -ForegroundColor Cyan
    }

}