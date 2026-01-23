<#
.Synopsis
Public Advisory Job Module

.DESCRIPTION
This script creates the job to process the Advisory data.

.Link
https://github.com/microsoft/ARI/Modules/Public/PublicFunctions/Jobs/Start-ARIAdvisoryJob.ps1

.COMPONENT
    This powershell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.9
First Release Date: 15th Oct, 2024
Authors: Claudio Merola

#>
function Start-ARIAdvisoryJob {
    param($Advisories)

    # Ensure Advisories is an array for safe iteration
    if ($null -eq $Advisories) {
        $Advisories = @()
    } elseif ($Advisories -isnot [System.Array]) {
        $Advisories = @($Advisories)
    }

    $tmp = foreach ($1 in $Advisories)
        {
            $data = $1.PROPERTIES

            # Handle advisories WITH resourceId (resource-level recommendations)
            if ($null -ne $data.resourceMetadata -and $null -ne $data.resourceMetadata.resourceId)
                {
                    # Safely access annualSavingsAmount
                    $Savings = 0
                    if ($null -ne $data.extendedProperties) {
                        try {
                            $hasAmount = $false
                            if ($data.extendedProperties -is [hashtable]) {
                                $hasAmount = $data.extendedProperties.ContainsKey('annualSavingsAmount')
                            } else {
                                $hasAmount = $data.extendedProperties.PSObject.Properties.Name -contains 'annualSavingsAmount'
                            }
                            if ($hasAmount -and -not [string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)) {
                                $Savings = $data.extendedProperties.annualSavingsAmount
                            }
                        } catch { }
                    }
                    
                    # Safely access savingsCurrency
                    $SavingsCurrency = 'USD'
                    if ($null -ne $data.extendedProperties) {
                        try {
                            $hasCurrency = $false
                            if ($data.extendedProperties -is [hashtable]) {
                                $hasCurrency = $data.extendedProperties.ContainsKey('savingsCurrency')
                            } else {
                                $hasCurrency = $data.extendedProperties.PSObject.Properties.Name -contains 'savingsCurrency'
                            }
                            if ($hasCurrency -and -not [string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)) {
                                $SavingsCurrency = $data.extendedProperties.savingsCurrency
                            }
                        } catch { }
                    }
                    $Resource = $data.resourceMetadata.resourceId.split('/')

                    # Safely extract resource information with bounds checking
                    $Subscription = ''
                    $ResourceGroup = ''
                    $ResourceType = ''
                    $ResourceName = ''
                    
                    if ($Resource.Count -lt 4) {
                        # Not enough segments - use impactedField/impactedValue
                        $ResourceType = $data.impactedField
                        $ResourceName = $data.impactedValue
                        # Try to get subscription if available
                        if ($Resource.Count -gt 2) {
                            $Subscription = $Resource[2]
                        }
                    }
                    elseif ($Resource.Count -lt 6) {
                        # Has subscription but not resource group
                        $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                        $ResourceType = $data.impactedField
                        $ResourceName = $data.impactedValue
                    }
                    elseif ($Resource.Count -lt 9) {
                        # Has subscription and resource group but not full resource path
                        $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                        $ResourceGroup = if ($Resource.Count -gt 4) { $Resource[4] } else { '' }
                        $ResourceType = $data.impactedField
                        $ResourceName = $data.impactedValue
                    }
                    else {
                        # Full resource path available
                        $Subscription = if ($Resource.Count -gt 2) { $Resource[2] } else { '' }
                        $ResourceGroup = if ($Resource.Count -gt 4) { $Resource[4] } else { '' }
                        $ResourceType = if ($Resource.Count -gt 7) { ($Resource[6] + '/' + $Resource[7]) } else { $data.impactedField }
                        $ResourceName = if ($Resource.Count -gt 8) { $Resource[8] } else { $data.impactedValue }
                    }

                    if ($data.impactedField -eq $ResourceType) {
                            $ImpactedField = ''
                    }
                    else {
                            $ImpactedField = $data.impactedField
                    }

                    if ($data.impactedValue -eq $ResourceName) {
                            $ImpactedValue = ''
                    }
                    else {
                            $ImpactedValue = $data.impactedValue
                        }

                    $obj = @{
                        'Subscription'           = $Subscription;
                        'Resource Group'         = $ResourceGroup;
                        'Resource Type'          = $ResourceType;
                        'Name'                   = $ResourceName;
                        'Detailed Type'          = $ImpactedField;
                        'Detailed Name'          = $ImpactedValue;
                        'Category'               = $data.category;
                        'Impact'                 = $data.impact;
                        'Description'            = if ($null -ne $data.shortDescription) { $data.shortDescription.problem } else { '' };
                        'SKU'                    = if ($null -ne $data.extendedProperties) { 
                            try { 
                                if ($data.extendedProperties -is [hashtable] -and $data.extendedProperties.ContainsKey('sku')) { $data.extendedProperties.sku } 
                                elseif ($data.extendedProperties.PSObject.Properties.Name -contains 'sku') { $data.extendedProperties.sku } 
                                else { '' } 
                            } catch { '' } 
                        } else { '' };
                        'Term'                   = if ($null -ne $data.extendedProperties) { 
                            try { 
                                if ($data.extendedProperties -is [hashtable] -and $data.extendedProperties.ContainsKey('term')) { $data.extendedProperties.term } 
                                elseif ($data.extendedProperties.PSObject.Properties.Name -contains 'term') { $data.extendedProperties.term } 
                                else { '' } 
                            } catch { '' } 
                        } else { '' };
                        'Look-back Period'       = if ($null -ne $data.extendedProperties) { 
                            try { 
                                if ($data.extendedProperties -is [hashtable] -and $data.extendedProperties.ContainsKey('lookbackPeriod')) { $data.extendedProperties.lookbackPeriod } 
                                elseif ($data.extendedProperties.PSObject.Properties.Name -contains 'lookbackPeriod') { $data.extendedProperties.lookbackPeriod } 
                                else { '' } 
                            } catch { '' } 
                        } else { '' };
                        'Quantity'               = if ($null -ne $data.extendedProperties) { 
                            try { 
                                if ($data.extendedProperties -is [hashtable] -and $data.extendedProperties.ContainsKey('qty')) { $data.extendedProperties.qty } 
                                elseif ($data.extendedProperties.PSObject.Properties.Name -contains 'qty') { $data.extendedProperties.qty } 
                                else { '' } 
                            } catch { '' } 
                        } else { '' };
                        'Savings Currency'       = $SavingsCurrency;
                        'Annual Savings'         = "=$Savings";
                        'Savings Region'         = if ($null -ne $data.extendedProperties) { 
                            try { 
                                if ($data.extendedProperties -is [hashtable] -and $data.extendedProperties.ContainsKey('region')) { $data.extendedProperties.region } 
                                elseif ($data.extendedProperties.PSObject.Properties.Name -contains 'region') { $data.extendedProperties.region } 
                                else { '' } 
                            } catch { '' } 
                        } else { '' }
                    }
                    $obj
                }
            # Handle advisories WITHOUT resourceId (subscription-level or management group-level recommendations)
            elseif ($null -ne $data)
                {
                    # Extract subscription ID from advisory ID if available
                    $Subscription = ''
                    if ($null -ne $1.id) {
                        # Advisory ID format: /subscriptions/{subId}/providers/Microsoft.Advisor/recommendations/{recId}
                        $idParts = $1.id -split '/'
                        $subIndex = [array]::IndexOf($idParts, 'subscriptions')
                        if ($subIndex -ge 0 -and $subIndex + 1 -lt $idParts.Count) {
                            $Subscription = $idParts[$subIndex + 1]
                        }
                    }
                    
                    # Use impactedField/impactedValue for resource type/name
                    $ResourceType = if ($null -ne $data.impactedField) { $data.impactedField } else { '' }
                    $ResourceName = if ($null -ne $data.impactedValue) { $data.impactedValue } else { '' }
                    
                    # Safely access annualSavingsAmount
                    $Savings = 0
                    if ($null -ne $data.extendedProperties) {
                        try {
                            $hasAmount = $false
                            if ($data.extendedProperties -is [hashtable]) {
                                $hasAmount = $data.extendedProperties.ContainsKey('annualSavingsAmount')
                            } else {
                                $hasAmount = $data.extendedProperties.PSObject.Properties.Name -contains 'annualSavingsAmount'
                            }
                            if ($hasAmount -and -not [string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)) {
                                $Savings = $data.extendedProperties.annualSavingsAmount
                            }
                        } catch { }
                    }
                    
                    # Safely access savingsCurrency
                    $SavingsCurrency = 'USD'
                    if ($null -ne $data.extendedProperties) {
                        try {
                            $hasCurrency = $false
                            if ($data.extendedProperties -is [hashtable]) {
                                $hasCurrency = $data.extendedProperties.ContainsKey('savingsCurrency')
                            } else {
                                $hasCurrency = $data.extendedProperties.PSObject.Properties.Name -contains 'savingsCurrency'
                            }
                            if ($hasCurrency -and -not [string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)) {
                                $SavingsCurrency = $data.extendedProperties.savingsCurrency
                            }
                        } catch { }
                    }

                    $obj = @{
                        'Subscription'           = $Subscription;
                        'Resource Group'         = '';
                        'Resource Type'          = $ResourceType;
                        'Name'                   = $ResourceName;
                        'Detailed Type'          = '';
                        'Detailed Name'          = '';
                        'Category'               = if ($null -ne $data.category) { $data.category } else { '' };
                        'Impact'                 = if ($null -ne $data.impact) { $data.impact } else { '' };
                        'Description'            = if ($null -ne $data.shortDescription) { $data.shortDescription.problem } else { '' };
                        'SKU'                    = if ($null -ne $data.extendedProperties) { 
                            try { 
                                if ($data.extendedProperties -is [hashtable] -and $data.extendedProperties.ContainsKey('sku')) { $data.extendedProperties.sku } 
                                elseif ($data.extendedProperties.PSObject.Properties.Name -contains 'sku') { $data.extendedProperties.sku } 
                                else { '' } 
                            } catch { '' } 
                        } else { '' };
                        'Term'                   = if ($null -ne $data.extendedProperties) { 
                            try { 
                                if ($data.extendedProperties -is [hashtable] -and $data.extendedProperties.ContainsKey('term')) { $data.extendedProperties.term } 
                                elseif ($data.extendedProperties.PSObject.Properties.Name -contains 'term') { $data.extendedProperties.term } 
                                else { '' } 
                            } catch { '' } 
                        } else { '' };
                        'Look-back Period'       = if ($null -ne $data.extendedProperties) { 
                            try { 
                                if ($data.extendedProperties -is [hashtable] -and $data.extendedProperties.ContainsKey('lookbackPeriod')) { $data.extendedProperties.lookbackPeriod } 
                                elseif ($data.extendedProperties.PSObject.Properties.Name -contains 'lookbackPeriod') { $data.extendedProperties.lookbackPeriod } 
                                else { '' } 
                            } catch { '' } 
                        } else { '' };
                        'Quantity'               = if ($null -ne $data.extendedProperties) { 
                            try { 
                                if ($data.extendedProperties -is [hashtable] -and $data.extendedProperties.ContainsKey('qty')) { $data.extendedProperties.qty } 
                                elseif ($data.extendedProperties.PSObject.Properties.Name -contains 'qty') { $data.extendedProperties.qty } 
                                else { '' } 
                            } catch { '' } 
                        } else { '' };
                        'Savings Currency'       = $SavingsCurrency;
                        'Annual Savings'         = "=$Savings";
                        'Savings Region'         = if ($null -ne $data.extendedProperties) { 
                            try { 
                                if ($data.extendedProperties -is [hashtable] -and $data.extendedProperties.ContainsKey('region')) { $data.extendedProperties.region } 
                                elseif ($data.extendedProperties.PSObject.Properties.Name -contains 'region') { $data.extendedProperties.region } 
                                else { '' } 
                            } catch { '' } 
                        } else { '' }
                    }
                    $obj
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

