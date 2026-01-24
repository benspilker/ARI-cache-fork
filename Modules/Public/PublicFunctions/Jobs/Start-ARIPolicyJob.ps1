<#
.Synopsis
Start Policy Job Module

.DESCRIPTION
This script processes and creates the Policy sheet based on advisor resources.

.Link
https://github.com/microsoft/ARI/Modules/Public/PublicFunctions/Jobs/Start-ARIPolicyJob.ps1

.COMPONENT
    This powershell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola

#>
function Start-ARIPolicyJob {
    param($Subscriptions, $PolicySetDef, $PolicyAssign, $PolicyDef)

    # Ensure PolicyDef is an array for safe iteration
    if ($null -eq $PolicyDef) {
        $PolicyDef = @()
    } elseif ($PolicyDef -isnot [System.Array]) {
        $PolicyDef = @($PolicyDef)
    }

    $poltmp = $PolicyDef | Select-Object -Property id,properties -Unique

    # Safely access PolicyAssign.policyAssignments
    $policyAssignments = @()
    if ($null -ne $PolicyAssign) {
        if ($PolicyAssign -is [PSCustomObject] -or $PolicyAssign -is [System.Collections.Hashtable]) {
            if ($null -ne $PolicyAssign.policyAssignments) {
                $policyAssignments = if ($PolicyAssign.policyAssignments -is [System.Array]) { 
                    $PolicyAssign.policyAssignments 
                } else { 
                    @($PolicyAssign.policyAssignments) 
                }
            }
        } elseif ($PolicyAssign -is [System.Array]) {
            $policyAssignments = $PolicyAssign
        }
    }

    # Ensure PolicySetDef is an array for safe iteration
    if ($null -eq $PolicySetDef) {
        $PolicySetDef = @()
    } elseif ($PolicySetDef -isnot [System.Array]) {
        $PolicySetDef = @($PolicySetDef)
    }

    $tmp = foreach ($1 in $policyAssignments)
        {
            # Safely check if policySetDefinitionId property exists and is not null/empty
            $policySetDefId = $null
            if ($null -ne $1) {
                try {
                    if ($1 -is [PSCustomObject]) {
                        # Check if property exists using PSObject.Properties
                        if ($1.PSObject.Properties['policySetDefinitionId']) {
                            $policySetDefId = $1.policySetDefinitionId
                        }
                    } elseif ($1 -is [System.Collections.Hashtable] -or $1 -is [System.Collections.IDictionary]) {
                        if ($1.ContainsKey('policySetDefinitionId')) {
                            $policySetDefId = $1['policySetDefinitionId']
                        }
                    } else {
                        # Try direct access as fallback
                        $policySetDefId = $1.policySetDefinitionId
                    }
                } catch {
                    # Property doesn't exist or access failed
                    $policySetDefId = $null
                }
            }
            
            if(![string]::IsNullOrEmpty($policySetDefId))
                {
                    $TempPolDef = foreach ($PolDe in $PolicySetDef)
                        {
                            if ($PolDe.id -eq $policySetDefId)
                                {
                                    # Safely access properties.displayName - handle cases where properties might not exist
                                    $displayName = $null
                                    if ($null -ne $PolDe) {
                                        try {
                                            if ($PolDe -is [PSCustomObject] -and $PolDe.PSObject.Properties['properties']) {
                                                $displayName = $PolDe.properties.displayName
                                            } elseif (($PolDe -is [System.Collections.Hashtable] -or $PolDe -is [System.Collections.IDictionary]) -and $PolDe.ContainsKey('properties')) {
                                                $displayName = $PolDe['properties'].displayName
                                            } elseif ($PolDe -is [PSCustomObject] -and $PolDe.PSObject.Properties['displayName']) {
                                                # Fallback: displayName might be at root level
                                                $displayName = $PolDe.displayName
                                            } elseif (($PolDe -is [System.Collections.Hashtable] -or $PolDe -is [System.Collections.IDictionary]) -and $PolDe.ContainsKey('displayName')) {
                                                # Fallback: displayName might be at root level
                                                $displayName = $PolDe['displayName']
                                            }
                                        } catch {
                                            # Property access failed - use empty string
                                            $displayName = ''
                                        }
                                    }
                                    $displayName
                                }
                        }
                    # Ensure TempPolDef is an array for safe count access
                    if ($null -eq $TempPolDef) {
                        $TempPolDef = @()
                    } elseif ($TempPolDef -isnot [System.Array]) {
                        $TempPolDef = @($TempPolDef)
                    }
                    $Initiative = if($TempPolDef.Count -gt 1){$TempPolDef[0]}else{$TempPolDef}
                    # Safely access results properties
                    $InitNonCompRes = ''
                    $InitNonCompPol = ''
                    if ($null -ne $1) {
                        try {
                            $resultsObj = $null
                            if ($1 -is [PSCustomObject] -and $1.PSObject.Properties['results']) {
                                $resultsObj = $1.results
                            } elseif (($1 -is [System.Collections.Hashtable] -or $1 -is [System.Collections.IDictionary]) -and $1.ContainsKey('results')) {
                                $resultsObj = $1['results']
                            }
                            if ($null -ne $resultsObj) {
                                if ($resultsObj -is [PSCustomObject] -and $resultsObj.PSObject.Properties['nonCompliantResources']) {
                                    $InitNonCompRes = $resultsObj.nonCompliantResources
                                } elseif (($resultsObj -is [System.Collections.Hashtable] -or $resultsObj -is [System.Collections.IDictionary]) -and $resultsObj.ContainsKey('nonCompliantResources')) {
                                    $InitNonCompRes = $resultsObj['nonCompliantResources']
                                }
                                if ($resultsObj -is [PSCustomObject] -and $resultsObj.PSObject.Properties['nonCompliantPolicies']) {
                                    $InitNonCompPol = $resultsObj.nonCompliantPolicies
                                } elseif (($resultsObj -is [System.Collections.Hashtable] -or $resultsObj -is [System.Collections.IDictionary]) -and $resultsObj.ContainsKey('nonCompliantPolicies')) {
                                    $InitNonCompPol = $resultsObj['nonCompliantPolicies']
                                }
                            }
                        } catch {
                            # Property access failed - use empty strings
                            $InitNonCompRes = ''
                            $InitNonCompPol = ''
                        }
                    }
                }
            else
                {
                    $Initiative = ''
                    $InitNonCompRes = ''
                    $InitNonCompPol = ''
                }

            # Safely access policyDefinitions
            $policyDefinitions = @()
            if ($null -ne $1) {
                try {
                    if ($1 -is [PSCustomObject]) {
                        # Check if property exists using PSObject.Properties
                        if ($1.PSObject.Properties['policyDefinitions']) {
                            $policyDefsValue = $1.policyDefinitions
                            if ($null -ne $policyDefsValue) {
                                $policyDefinitions = if ($policyDefsValue -is [System.Array]) { 
                                    $policyDefsValue 
                                } else { 
                                    @($policyDefsValue) 
                                }
                            }
                        }
                    } elseif ($1 -is [System.Collections.Hashtable] -or $1 -is [System.Collections.IDictionary]) {
                        if ($1.ContainsKey('policyDefinitions')) {
                            $policyDefsValue = $1['policyDefinitions']
                            if ($null -ne $policyDefsValue) {
                                $policyDefinitions = if ($policyDefsValue -is [System.Array]) { 
                                    $policyDefsValue 
                                } else { 
                                    @($policyDefsValue) 
                                }
                            }
                        }
                    }
                } catch {
                    # Property doesn't exist or access failed - policyDefinitions remains empty array
                    $policyDefinitions = @()
                }
            }

            foreach ($2 in $policyDefinitions)
                {
                    # Safely access policyDefinitionId
                    $policyDefId = $null
                    if ($null -ne $2) {
                        try {
                            if ($2 -is [PSCustomObject] -and $2.PSObject.Properties['policyDefinitionId']) {
                                $policyDefId = $2.policyDefinitionId
                            } elseif (($2 -is [System.Collections.Hashtable] -or $2 -is [System.Collections.IDictionary]) -and $2.ContainsKey('policyDefinitionId')) {
                                $policyDefId = $2['policyDefinitionId']
                            }
                        } catch {
                            $policyDefId = $null
                        }
                    }
                    
                    # Safely access PolicyDef properties from poltmp
                    $Pol = $null
                    if ($null -ne $policyDefId) {
                        $matchingPolDef = $poltmp | Where-Object {$_.id -eq $policyDefId} | Select-Object -First 1
                        if ($null -ne $matchingPolDef) {
                            try {
                                # Check if properties exists and is not null
                                if ($matchingPolDef -is [PSCustomObject] -and $matchingPolDef.PSObject.Properties['properties'] -and $null -ne $matchingPolDef.properties) {
                                    $Pol = $matchingPolDef.properties
                                } elseif (($matchingPolDef -is [System.Collections.Hashtable] -or $matchingPolDef -is [System.Collections.IDictionary]) -and $matchingPolDef.ContainsKey('properties') -and $null -ne $matchingPolDef['properties']) {
                                    $Pol = $matchingPolDef['properties']
                                }
                            } catch {
                                # Property access failed
                                $Pol = $null
                            }
                        }
                    }
                    
                    if(![string]::IsNullOrEmpty($Pol))
                        {
                            # Safely access results.resourceDetails
                            $resourceDetails = @()
                            if ($null -ne $2) {
                                try {
                                    $resultsObj = $null
                                    if ($2 -is [PSCustomObject] -and $2.PSObject.Properties['results']) {
                                        $resultsObj = $2.results
                                    } elseif (($2 -is [System.Collections.Hashtable] -or $2 -is [System.Collections.IDictionary]) -and $2.ContainsKey('results')) {
                                        $resultsObj = $2['results']
                                    }
                                    if ($null -ne $resultsObj) {
                                        if ($resultsObj -is [PSCustomObject] -and $resultsObj.PSObject.Properties['resourceDetails']) {
                                            $resourceDetailsValue = $resultsObj.resourceDetails
                                            if ($null -ne $resourceDetailsValue) {
                                                $resourceDetails = if ($resourceDetailsValue -is [System.Array]) { $resourceDetailsValue } else { @($resourceDetailsValue) }
                                            }
                                        } elseif (($resultsObj -is [System.Collections.Hashtable] -or $resultsObj -is [System.Collections.IDictionary]) -and $resultsObj.ContainsKey('resourceDetails')) {
                                            $resourceDetailsValue = $resultsObj['resourceDetails']
                                            if ($null -ne $resourceDetailsValue) {
                                                $resourceDetails = if ($resourceDetailsValue -is [System.Array]) { $resourceDetailsValue } else { @($resourceDetailsValue) }
                                            }
                                        }
                                    }
                                } catch {
                                    $resourceDetails = @()
                                }
                            }
                            
                            $PolResUnkown = ($resourceDetails | Where-Object {$_.complianceState -eq 'unknown'} | Select-Object -ExpandProperty Count)
                            $PolResUnkown = if (![string]::IsNullOrEmpty($PolResUnkown)){$PolResUnkown}else{'0'}
                            $PolResCompl = ($resourceDetails | Where-Object {$_.complianceState -eq 'compliant'} | Select-Object -ExpandProperty Count)
                            $PolResCompl = if (![string]::IsNullOrEmpty($PolResCompl)){$PolResCompl}else{'0'}
                            $PolResNonCompl = ($resourceDetails | Where-Object {$_.complianceState -eq 'noncompliant'} | Select-Object -ExpandProperty Count)
                            $PolResNonCompl = if (![string]::IsNullOrEmpty($PolResNonCompl)){$PolResNonCompl}else{'0'}
                            $PolResExemp = ($resourceDetails | Where-Object {$_.complianceState -eq 'exempt'} | Select-Object -ExpandProperty Count)
                            $PolResExemp = if (![string]::IsNullOrEmpty($PolResExemp)){$PolResExemp}else{'0'}

                            # Safely access $2.effect
                            $effectValue = ''
                            if ($null -ne $2) {
                                try {
                                    if ($2 -is [PSCustomObject] -and $2.PSObject.Properties['effect']) {
                                        $effectValue = $2.effect
                                    } elseif (($2 -is [System.Collections.Hashtable] -or $2 -is [System.Collections.IDictionary]) -and $2.ContainsKey('effect')) {
                                        $effectValue = $2['effect']
                                    }
                                } catch {
                                    $effectValue = ''
                                }
                            }
                            
                            # Safely access Pol properties
                            $polDisplayName = ''
                            $polPolicyType = ''
                            $polMode = ''
                            $polVersion = ''
                            $polDeprecated = ''
                            $polCategory = ''
                            
                            if ($null -ne $Pol) {
                                try {
                                    if ($Pol -is [PSCustomObject]) {
                                        $polDisplayName = if ($Pol.PSObject.Properties['displayName']) { $Pol.displayName } else { '' }
                                        $polPolicyType = if ($Pol.PSObject.Properties['policyType']) { $Pol.policyType } else { '' }
                                        $polMode = if ($Pol.PSObject.Properties['mode']) { $Pol.mode } else { '' }
                                        $polVersion = if ($Pol.PSObject.Properties['version']) { $Pol.version } else { '' }
                                        if ($Pol.PSObject.Properties['metadata'] -and $null -ne $Pol.metadata) {
                                            $polDeprecated = if ($Pol.metadata.PSObject.Properties['deprecated']) { $Pol.metadata.deprecated } else { '' }
                                            $polCategory = if ($Pol.metadata.PSObject.Properties['category']) { $Pol.metadata.category } else { '' }
                                        }
                                    } elseif ($Pol -is [System.Collections.Hashtable] -or $Pol -is [System.Collections.IDictionary]) {
                                        $polDisplayName = if ($Pol.ContainsKey('displayName')) { $Pol['displayName'] } else { '' }
                                        $polPolicyType = if ($Pol.ContainsKey('policyType')) { $Pol['policyType'] } else { '' }
                                        $polMode = if ($Pol.ContainsKey('mode')) { $Pol['mode'] } else { '' }
                                        $polVersion = if ($Pol.ContainsKey('version')) { $Pol['version'] } else { '' }
                                        if ($Pol.ContainsKey('metadata') -and $null -ne $Pol['metadata']) {
                                            $metadata = $Pol['metadata']
                                            if ($metadata -is [System.Collections.Hashtable] -or $metadata -is [System.Collections.IDictionary]) {
                                                $polDeprecated = if ($metadata.ContainsKey('deprecated')) { $metadata['deprecated'] } else { '' }
                                                $polCategory = if ($metadata.ContainsKey('category')) { $metadata['category'] } else { '' }
                                            }
                                        }
                                    }
                                } catch {
                                    # Property access failed - use empty strings
                                }
                            }
                            
                            $obj = @{
                                'Initiative'                            = $Initiative;
                                'Initiative Non Compliance Resources'   = $InitNonCompRes;
                                'Initiative Non Compliance Policies'    = $InitNonCompPol;
                                'Policy'                                = $polDisplayName;
                                'Policy Type'                           = $polPolicyType;
                                'Effect'                                = $effectValue;
                                'Compliance Resources'                  = $PolResCompl;
                                'Non Compliance Resources'              = $PolResNonCompl;
                                'Unknown Resources'                     = $PolResUnkown;
                                'Exempt Resources'                      = $PolResExemp
                                'Policy Mode'                           = $polMode;
                                'Policy Version'                        = $polVersion;
                                'Policy Deprecated'                     = $polDeprecated;
                                'Policy Category'                       = $polCategory
                            }
                            $obj
                        }
                }
        }
    # Ensure tmp is always an array (foreach might return null if empty)
    if ($null -eq $tmp) {
        $tmp = @()
    } elseif ($tmp -isnot [System.Array]) {
        $tmp = @($tmp)
    }
    $tmp
}