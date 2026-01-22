<#
.Synopsis
Module for Security Center Report

.DESCRIPTION
This script processes and creates the Security Center sheet in the Excel report.

.Link
https://github.com/microsoft/ARI/Modules/Private/3.ReportingFunctions/Build-ARISecCenterReport.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Build-ARISecCenterReport {
    param($File, $Sec, $TableStyle)
    
    # Ensure Sec is an array for safe handling
    if ($null -eq $Sec) {
        $Sec = @()
    } elseif ($Sec -isnot [System.Array]) {
        $Sec = @($Sec)
    }
    
    # Only create sheet if we have data
    if ($Sec.Count -eq 0) {
        Write-Debug "  No security center data to report - skipping Security Center sheet"
        return
    }
    
    $condtxtsec = $(New-ConditionalText High -Range G:G
    New-ConditionalText High -Range L:L)

    [PSCustomObject]$Sec |
    ForEach-Object { $_ } |
    Select-Object 'Subscription',
    'Resource Group',
    'Resource Type',
    'Resource Name',
    'Categories',
    'Control',
    'Severity',
    'Status',
    'Remediation',
    'Remediation Effort',
    'User Impact',
    'Threats' |
    Export-Excel -Path $File -WorksheetName 'SecurityCenter' -AutoSize -MaxAutoSizeRows 100 -MoveToStart -TableName 'SecurityCenter' -TableStyle $tableStyle -ConditionalText $condtxtsec
}