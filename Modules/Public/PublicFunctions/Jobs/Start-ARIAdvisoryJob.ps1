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
                    $Savings = if([string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){0}Else{$data.extendedProperties.annualSavingsAmount}
                    $SavingsCurrency = if([string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){'USD'}Else{$data.extendedProperties.savingsCurrency}
                    $Resource = $data.resourceMetadata.resourceId.split('/')

                    if ($Resource.Count -lt 4) {
                        $ResourceType = $data.impactedField
                        $ResourceName = $data.impactedValue
                    }
                    else {
                        $ResourceType = ($Resource[6] + '/' + $Resource[7])
                        $ResourceName = $Resource[8]
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
                        'Subscription'           = $Resource[2];
                        'Resource Group'         = $Resource[4];
                        'Resource Type'          = $ResourceType;
                        'Name'                   = $ResourceName;
                        'Detailed Type'          = $ImpactedField;
                        'Detailed Name'          = $ImpactedValue;
                        'Category'               = $data.category;
                        'Impact'                 = $data.impact;
                        'Description'            = if ($null -ne $data.shortDescription) { $data.shortDescription.problem } else { '' };
                        'SKU'                    = $data.extendedProperties.sku;
                        'Term'                   = $data.extendedProperties.term;
                        'Look-back Period'       = $data.extendedProperties.lookbackPeriod;
                        'Quantity'               = $data.extendedProperties.qty;
                        'Savings Currency'       = $SavingsCurrency;
                        'Annual Savings'         = "=$Savings";
                        'Savings Region'         = $data.extendedProperties.region
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
                    
                    $Savings = if([string]::IsNullOrEmpty($data.extendedProperties.annualSavingsAmount)){0}Else{$data.extendedProperties.annualSavingsAmount}
                    $SavingsCurrency = if([string]::IsNullOrEmpty($data.extendedProperties.savingsCurrency)){'USD'}Else{$data.extendedProperties.savingsCurrency}

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
                        'SKU'                    = if ($null -ne $data.extendedProperties) { $data.extendedProperties.sku } else { '' };
                        'Term'                   = if ($null -ne $data.extendedProperties) { $data.extendedProperties.term } else { '' };
                        'Look-back Period'       = if ($null -ne $data.extendedProperties) { $data.extendedProperties.lookbackPeriod } else { '' };
                        'Quantity'               = if ($null -ne $data.extendedProperties) { $data.extendedProperties.qty } else { '' };
                        'Savings Currency'       = $SavingsCurrency;
                        'Annual Savings'         = "=$Savings";
                        'Savings Region'         = if ($null -ne $data.extendedProperties) { $data.extendedProperties.region } else { '' }
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

