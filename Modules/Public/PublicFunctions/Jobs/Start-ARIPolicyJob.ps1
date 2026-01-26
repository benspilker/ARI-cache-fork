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

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Starting with PolicyDef count: ' + $(if ($null -eq $PolicyDef) { 'null' } elseif ($PolicyDef -is [System.Array]) { $PolicyDef.Count } else { '1 (not array)' }))
    
    # Ensure PolicyDef is an array for safe iteration
    if ($null -eq $PolicyDef) {
        $PolicyDef = @()
    } elseif ($PolicyDef -isnot [System.Array]) {
        $PolicyDef = @($PolicyDef)
    }
    
    # Debug: Show sample PolicyDef structure
    if ($PolicyDef.Count -gt 0) {
        $samplePolDef = $PolicyDef[0]
        $sampleId = $null
        if ($samplePolDef -is [PSCustomObject] -and $samplePolDef.PSObject.Properties['id']) {
            $sampleId = $samplePolDef.id
        } elseif (($samplePolDef -is [System.Collections.Hashtable] -or $samplePolDef -is [System.Collections.IDictionary]) -and $samplePolDef.ContainsKey('id')) {
            $sampleId = $samplePolDef['id']
        }
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Sample PolicyDef ID format: ' + $(if ($sampleId) { $sampleId } else { 'NO_ID_FOUND' }))
    }

    # Create poltmp - handle cases where properties might be at root level or nested
    $poltmp = @()
    $poltmpCount = 0
    foreach ($polDefItem in $PolicyDef) {
        if ($null -eq $polDefItem) { continue }
        
        $polId = $null
        $polProperties = $null
        
        # Get id
        if ($polDefItem -is [PSCustomObject] -and $polDefItem.PSObject.Properties['id']) {
            $polId = $polDefItem.id
        } elseif (($polDefItem -is [System.Collections.Hashtable] -or $polDefItem -is [System.Collections.IDictionary]) -and $polDefItem.ContainsKey('id')) {
            $polId = $polDefItem['id']
        }
        
        # Get properties - check nested first, then root level
        if ($polDefItem -is [PSCustomObject]) {
            if ($polDefItem.PSObject.Properties['properties'] -and $null -ne $polDefItem.properties) {
                $polProperties = $polDefItem.properties
            } else {
                # Properties might be at root level - use the entire object as properties
                $polProperties = $polDefItem
            }
        } elseif ($polDefItem -is [System.Collections.Hashtable] -or $polDefItem -is [System.Collections.IDictionary]) {
            if ($polDefItem.ContainsKey('properties') -and $null -ne $polDefItem['properties']) {
                $polProperties = $polDefItem['properties']
            } else {
                # Properties might be at root level - use the entire object as properties
                $polProperties = $polDefItem
            }
        }
        
        if ($null -ne $polId -and $null -ne $polProperties) {
            $poltmp += [PSCustomObject]@{
                id = $polId
                properties = $polProperties
            }
            $poltmpCount++
        }
    }
    
    # Remove duplicates by id while preserving full objects
    $poltmp = $poltmp | Sort-Object -Property id -Unique
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Created poltmp with ' + $poltmp.Count + ' unique PolicyDef entries (processed ' + $poltmpCount + ' items)')
    
    # Debug: Show sample of PolicyDef IDs to help diagnose matching issues
    # Check for management group vs subscription level PolicyDefs
    $mgPolicyDefs = $poltmp | Where-Object { $_.id -match 'managementgroups' }
    $subPolicyDefs = $poltmp | Where-Object { $_.id -notmatch 'managementgroups' }
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: PolicyDef breakdown - MG level: ' + $mgPolicyDefs.Count + ', Subscription level: ' + $subPolicyDefs.Count)
    if ($mgPolicyDefs.Count -gt 0) {
        $sampleMgIds = $mgPolicyDefs | Select-Object -First 3 | ForEach-Object { $_.id }
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Sample MG PolicyDef IDs: ' + ($sampleMgIds -join ' | '))
    }
    if ($poltmp.Count -gt 0) {
        $sampleIds = $poltmp | Select-Object -First 5 | ForEach-Object { $_.id }
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Sample PolicyDef IDs (first 5): ' + ($sampleIds -join ' | '))
        
        # Also show GUIDs extracted from sample IDs
        $sampleGuids = $sampleIds | ForEach-Object {
            if ($_ -match 'policydefinitions/([a-f0-9\-]{36})$') {
                $Matches[1]
            } elseif ($_ -match '^[a-f0-9\-]{36}$') {
                $_
            } else {
                'NO_GUID'
            }
        }
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Sample PolicyDef GUIDs (first 5): ' + ($sampleGuids -join ' | '))
    } else {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: WARNING - poltmp is empty! PolicyDef count: ' + $PolicyDef.Count)
    }

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

    $tmp = @()
    $processedAssignments = 0
    $processedPolicyDefs = 0
    foreach ($1 in $policyAssignments)
        {
            $processedAssignments++
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
                    # Extract GUID from policySetDefinitionId for fallback matching
                    $policySetDefGuid = $null
                    if ($policySetDefId -match 'policysetdefinitions/([a-fA-F0-9\-]{36})$') {
                        $policySetDefGuid = $Matches[1].ToLower()
                    } elseif ($policySetDefId -match '^[a-fA-F0-9\-]{36}$') {
                        $policySetDefGuid = $policySetDefId.ToLower()
                    }
                    
                    $TempPolDef = foreach ($PolDe in $PolicySetDef)
                        {
                            $polSetDefId = $null
                            if ($PolDe -is [PSCustomObject] -and $PolDe.PSObject.Properties['id']) {
                                $polSetDefId = $PolDe.id
                            } elseif (($PolDe -is [System.Collections.Hashtable] -or $PolDe -is [System.Collections.IDictionary]) -and $PolDe.ContainsKey('id')) {
                                $polSetDefId = $PolDe['id']
                            }
                            
                            $isMatch = $false
                            if ($null -ne $polSetDefId) {
                                # Try case-insensitive exact match first
                                if ($polSetDefId -eq $policySetDefId -or $polSetDefId.ToLower() -eq $policySetDefId.ToLower()) {
                                    $isMatch = $true
                                }
                                # Try GUID match if exact match failed
                                elseif ($null -ne $policySetDefGuid) {
                                    $polSetDefGuid = $null
                                    if ($polSetDefId -match 'policysetdefinitions/([a-fA-F0-9\-]{36})$') {
                                        $polSetDefGuid = $Matches[1].ToLower()
                                    } elseif ($polSetDefId -match '^[a-fA-F0-9\-]{36}$') {
                                        $polSetDefGuid = $polSetDefId.ToLower()
                                    }
                                    if ($null -ne $polSetDefGuid -and $polSetDefGuid -eq $policySetDefGuid) {
                                        $isMatch = $true
                                    }
                                }
                            }
                            
                            if ($isMatch)
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
            
            $policyDefsInAssignment = $policyDefinitions.Count
            Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Assignment ' + $processedAssignments + ' has ' + $policyDefsInAssignment + ' policy definition(s)')

            foreach ($2 in $policyDefinitions)
                {
                    # Safely access policyDefinitionId (policyDefinitions may be objects or raw strings)
                    $policyDefId = $null
                    if ($null -ne $2) {
                        try {
                            if ($2 -is [string]) {
                                $policyDefId = $2
                            } elseif ($2 -is [PSCustomObject] -and $2.PSObject.Properties['policyDefinitionId']) {
                                $policyDefId = $2.policyDefinitionId
                            } elseif (($2 -is [System.Collections.Hashtable] -or $2 -is [System.Collections.IDictionary]) -and $2.ContainsKey('policyDefinitionId')) {
                                $policyDefId = $2['policyDefinitionId']
                            }
                        } catch {
                            $policyDefId = $null
                        }
                    }
                    
                    # Safely access PolicyDef properties from poltmp
                    # Match policyDefinitionId (which may be full path or GUID) against PolicyDef id (which may also be full path or GUID)
                    $Pol = $null
                    if ($null -ne $policyDefId) {
                        # Extract GUID from policyDefinitionId if it's a full path (case-insensitive)
                        # Format: /providers/microsoft.authorization/policydefinitions/{guid} or /providers/Microsoft.Authorization/policyDefinitions/{guid}
                        # or: /providers/microsoft.management/managementgroups/{mg}/providers/microsoft.authorization/policydefinitions/{guid}
                        $policyDefGuid = $null
                        # Case-insensitive regex match for GUID extraction
                        if ($policyDefId -match 'policydefinitions/([a-fA-F0-9\-]{36})') {
                            $policyDefGuid = $Matches[1].ToLower()  # Normalize to lowercase for comparison
                        } elseif ($policyDefId -match '^[a-fA-F0-9\-]{36}$') {
                            # Already a GUID - normalize to lowercase
                            $policyDefGuid = $policyDefId.ToLower()
                        }
                        
                        # Try to find matching PolicyDef by:
                        # 1. Case-insensitive exact match (full path to full path)
                        # 2. Case-insensitive GUID match (extract GUID from PolicyDef id and compare)
                        $matchingPolDef = $null
                        foreach ($polDefItem in $poltmp) {
                            $polDefId = $null
                            if ($polDefItem -is [PSCustomObject] -and $polDefItem.PSObject.Properties['id']) {
                                $polDefId = $polDefItem.id
                            } elseif (($polDefItem -is [System.Collections.Hashtable] -or $polDefItem -is [System.Collections.IDictionary]) -and $polDefItem.ContainsKey('id')) {
                                $polDefId = $polDefItem['id']
                            }
                            
                            if ($null -ne $polDefId) {
                                # Try case-insensitive exact match first
                                if ($polDefId -eq $policyDefId -or $polDefId.ToLower() -eq $policyDefId.ToLower()) {
                                    $matchingPolDef = $polDefItem
                                    break
                                }
                                
                                # Try GUID match if we extracted a GUID (case-insensitive)
                                if ($null -ne $policyDefGuid) {
                                    $polDefGuid = $null
                                    # Case-insensitive regex match for GUID extraction
                                    if ($polDefId -match 'policydefinitions/([a-fA-F0-9\-]{36})') {
                                        $polDefGuid = $Matches[1].ToLower()  # Normalize to lowercase
                                    } elseif ($polDefId -match '^[a-fA-F0-9\-]{36}$') {
                                        $polDefGuid = $polDefId.ToLower()  # Normalize to lowercase
                                    }
                                    
                                    # Case-insensitive GUID comparison
                                    if ($null -ne $polDefGuid -and $polDefGuid -eq $policyDefGuid) {
                                        $matchingPolDef = $polDefItem
                                        break
                                    }
                                }
                            }
                        }
                        
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
                        } else {
                            # Try one more time: Look for PolicyDef with same name/GUID but different path
                            # This handles cases where assignment references MG-scoped PolicyDef but we have tenant-level one
                            $fallbackPolDef = $null
                            if ($null -ne $policyDefGuid) {
                                # Try to find any PolicyDef with the same GUID, regardless of path
                                foreach ($polDefItem in $poltmp) {
                                    $polDefId = $null
                                    if ($polDefItem -is [PSCustomObject] -and $polDefItem.PSObject.Properties['id']) {
                                        $polDefId = $polDefItem.id
                                    } elseif (($polDefItem -is [System.Collections.Hashtable] -or $polDefItem -is [System.Collections.IDictionary]) -and $polDefItem.ContainsKey('id')) {
                                        $polDefId = $polDefItem['id']
                                    }
                                    
                                    if ($null -ne $polDefId) {
                                        $polDefGuidCheck = $null
                                        if ($polDefId -match 'policydefinitions/([a-fA-F0-9\-]{36})') {
                                            $polDefGuidCheck = $Matches[1].ToLower()
                                        } elseif ($polDefId -match '^[a-fA-F0-9\-]{36}$') {
                                            $polDefGuidCheck = $polDefId.ToLower()
                                        }
                                        
                                        if ($null -ne $polDefGuidCheck -and $polDefGuidCheck -eq $policyDefGuid) {
                                            $fallbackPolDef = $polDefItem
                                            break
                                        }
                                    }
                                }
                            }
                            
                            if ($null -ne $fallbackPolDef) {
                                # Found a PolicyDef with matching GUID but different path - use it
                                try {
                                    if ($fallbackPolDef -is [PSCustomObject] -and $fallbackPolDef.PSObject.Properties['properties'] -and $null -ne $fallbackPolDef.properties) {
                                        $Pol = $fallbackPolDef.properties
                                    } elseif (($fallbackPolDef -is [System.Collections.Hashtable] -or $fallbackPolDef -is [System.Collections.IDictionary]) -and $fallbackPolDef.ContainsKey('properties') -and $null -ne $fallbackPolDef['properties']) {
                                        $Pol = $fallbackPolDef['properties']
                                    }
                                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Found PolicyDef by GUID fallback (different path): ' + $policyDefId)
                                } catch {
                                    $Pol = $null
                                }
                            }
                            
                            if ($null -eq $Pol) {
                                # More detailed debug: show what we tried to match against
                                $debugSampleIds = $poltmp | Select-Object -First 3 | ForEach-Object { $_.id }
                                $debugSampleGuids = $debugSampleIds | ForEach-Object {
                                    if ($_ -match 'policydefinitions/([a-fA-F0-9\-]{36})$') {
                                        $Matches[1].ToLower()
                                    } elseif ($_ -match '^[a-fA-F0-9\-]{36}$') {
                                        $_.ToLower()
                                    } else {
                                        'NO_GUID'
                                    }
                                }
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: No matching PolicyDef found for policyDefinitionId: ' + $policyDefId + ' (extracted GUID: ' + $(if ($policyDefGuid) { $policyDefGuid } else { 'none' }) + ')')
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Searched in ' + $poltmp.Count + ' PolicyDef(s). Sample IDs: ' + ($debugSampleIds -join ' | ') + ' | Sample GUIDs: ' + ($debugSampleGuids -join ' | '))
                                
                                # Last resort: Create a minimal PolicyDef object from the policyDefinitionId
                                # This will be filtered out by our GUID filtering logic since it has no displayName
                                $Pol = [PSCustomObject]@{
                                    displayName = if ($policyDefId -match '/([^/]+)$') { $Matches[1] } else { 'Unknown Policy' }
                                    policyType = if ($policyDefId -match '/managementgroups/') { 'Custom' } else { 'BuiltIn' }
                                    mode = 'All'
                                    version = ''
                                    metadata = @{
                                        deprecated = ''
                                        category = ''
                                    }
                                }
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Created minimal PolicyDef object for missing PolicyDef: ' + $policyDefId)
                            }
                        }
                    } else {
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: policyDefinitionId is null or empty')
                    }
                    
                    # Check if Pol is set (can be PSCustomObject or Hashtable, not just string)
                    if ($null -ne $Pol)
                        {
                            $processedPolicyDefs++
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
    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Start-ARIPolicyJob: Processed ' + $processedAssignments + ' assignment(s), ' + $processedPolicyDefs + ' policy definition(s), returning ' + $tmp.Count + ' Policy record(s)')
    $tmp
}
