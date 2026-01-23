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
                                    $PolDe.properties.displayName
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
                    $InitNonCompRes = if ($null -ne $1 -and $null -ne $1.results) { 
                        if ($null -ne $1.results.nonCompliantResources) { $1.results.nonCompliantResources } else { '' }
                    } else { '' }
                    $InitNonCompPol = if ($null -ne $1 -and $null -ne $1.results) { 
                        if ($null -ne $1.results.nonCompliantPolicies) { $1.results.nonCompliantPolicies } else { '' }
                    } else { '' }
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
                if ($1 -is [PSCustomObject] -or $1 -is [System.Collections.Hashtable]) {
                    if ($null -ne $1.policyDefinitions) {
                        $policyDefinitions = if ($1.policyDefinitions -is [System.Array]) { 
                            $1.policyDefinitions 
                        } else { 
                            @($1.policyDefinitions) 
                        }
                    }
                } elseif ($1 -is [System.Collections.IDictionary]) {
                    if ($1.ContainsKey('policyDefinitions')) {
                        $policyDefinitions = if ($1['policyDefinitions'] -is [System.Array]) { 
                            $1['policyDefinitions'] 
                        } else { 
                            @($1['policyDefinitions']) 
                        }
                    }
                }
            }

            foreach ($2 in $policyDefinitions)
                {
                    $Pol = (($poltmp | Where-Object {$_.id -eq $2.policyDefinitionId}).properties)
                    if(![string]::IsNullOrEmpty($Pol))
                        {
                            $PolResUnkown = ($2.results.resourceDetails | Where-Object {$_.complianceState -eq 'unknown'} | Select-Object -ExpandProperty Count)
                            $PolResUnkown = if (![string]::IsNullOrEmpty($PolResUnkown)){$PolResUnkown}else{'0'}
                            $PolResCompl = ($2.results.resourceDetails | Where-Object {$_.complianceState -eq 'compliant'} | Select-Object -ExpandProperty Count)
                            $PolResCompl = if (![string]::IsNullOrEmpty($PolResCompl)){$PolResCompl}else{'0'}
                            $PolResNonCompl = ($2.results.resourceDetails | Where-Object {$_.complianceState -eq 'noncompliant'} | Select-Object -ExpandProperty Count)
                            $PolResNonCompl = if (![string]::IsNullOrEmpty($PolResNonCompl)){$PolResNonCompl}else{'0'}
                            $PolResExemp = ($2.results.resourceDetails | Where-Object {$_.complianceState -eq 'exempt'} | Select-Object -ExpandProperty Count)
                            $PolResExemp = if (![string]::IsNullOrEmpty($PolResExemp)){$PolResExemp}else{'0'}

                            $obj = @{
                                'Initiative'                            = $Initiative;
                                'Initiative Non Compliance Resources'   = $InitNonCompRes;
                                'Initiative Non Compliance Policies'    = $InitNonCompPol;
                                'Policy'                                = $Pol.displayName;
                                'Policy Type'                           = $Pol.policyType;
                                'Effect'                                = $2.effect;
                                'Compliance Resources'                  = $PolResCompl;
                                'Non Compliance Resources'              = $PolResNonCompl;
                                'Unknown Resources'                     = $PolResUnkown;
                                'Exempt Resources'                      = $PolResExemp
                                'Policy Mode'                           = $Pol.mode;
                                'Policy Version'                        = $Pol.version;
                                'Policy Deprecated'                     = $Pol.metadata.deprecated;
                                'Policy Category'                       = $Pol.metadata.category
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