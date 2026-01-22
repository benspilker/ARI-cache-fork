<#
.Synopsis
Module for Advisory Report

.DESCRIPTION
This script processes and creates the Advisory sheet in the Excel report.

.Link
https://github.com/microsoft/ARI/Modules/Private/3.ReportingFunctions/Build-ARIAdvisoryReport.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.9
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Build-ARIAdvisoryReport {
    param($File, $Adv, $TableStyle)
    
    # Ensure Adv is an array for safe handling
    if ($null -eq $Adv) {
        $Adv = @()
    } elseif ($Adv -isnot [System.Array]) {
        $Adv = @($Adv)
    }
    
    # Only create sheet if we have data
    if ($Adv.Count -eq 0) {
        Write-Debug "  No advisory data to report - skipping Advisor sheet"
        return
    }
    
    $condtxtadv = @()
    $condtxtadv += New-ConditionalText High -Range H:H
    $condtxtadv += New-ConditionalText Security -Range G:G -BackgroundColor Wheat

    $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '#,##0.00' -Range O:O

    [PSCustomObject]$Adv |
    ForEach-Object { $_ } |
    Select-Object 'Subscription',
    'Resource Group',
    'Resource Type',
    'Name',
    'Detailed Type',
    'Detailed Name',
    'Category',
    'Impact',
    'Description',
    'SKU',
    'Term',
    'Look-back Period',
    'Quantity',
    'Savings Currency',
    'Annual Savings',
    'Savings Region' |
    Export-Excel -Path $File -WorksheetName 'Advisor' -AutoSize -MaxAutoSizeRows 100 -TableName 'AzureAdvisory' -MoveToStart -TableStyle $tableStyle -Style $Style -ConditionalText $condtxtadv
}