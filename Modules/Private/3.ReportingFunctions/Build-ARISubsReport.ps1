<#
.Synopsis
Module for Subscription Report

.DESCRIPTION
This script processes and creates the Subscription sheet in the Excel report.

.Link
https://github.com/microsoft/ARI/Modules/Private/3.ReportingFunctions/Build-ARISubsReport.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Build-ARISubsReport {
    param($File, $Sub, $IncludeCosts, $TableStyle)
    
    # Ensure $Sub is an array
    if ($null -eq $Sub) {
        $Sub = @()
    } elseif ($Sub -isnot [System.Array]) {
        $Sub = @($Sub)
    }
    
    # Safely get subscription count - handle null/empty cases
    $subscriptionCount = 0
    if ($Sub.Count -gt 0) {
        try {
            $subscriptionValues = $Sub | Select-Object -ExpandProperty 'Subscription' -ErrorAction SilentlyContinue | Where-Object { $null -ne $_ }
            if ($subscriptionValues) {
                $subscriptionCount = ($subscriptionValues | Select-Object -Unique).Count
            } else {
                # If no Subscription property, use array count
                $subscriptionCount = $Sub.Count
            }
        } catch {
            # If property doesn't exist, try to get count from array itself
            $subscriptionCount = $Sub.Count
        }
    }
    
    $TableName = ('SubsTable_'+$subscriptionCount)

    # Only create sheet if we have data
    if ($Sub.Count -eq 0) {
        Write-Debug "  No subscription data to report - skipping Subscriptions sheet"
        return
    }

    if ($IncludeCosts.IsPresent)
        {
            $Style = @() 
            $Style += New-ExcelStyle -AutoSize -HorizontalAlignment Center -NumberFormat '0'
            $Style += New-ExcelStyle -Width 55 -NumberFormat '$#,#########0.000000000' -Range J:J
            $Style += New-ExcelStyle -AutoSize -NumberFormat '$#,##0.00' -Range I:I
            [PSCustomObject]$Sub |
                ForEach-Object { $_ } |
                Select-Object 'Subscription',
                'Resource Group',
                'Location',
                'Resource Type',
                'Service Name',
                'Currency',
                'Month',
                'Year',
                'Cost',
                'Detailed Cost' | Export-Excel -Path $File -WorksheetName 'Subscriptions' -TableName $TableName -TableStyle $TableStyle -Style $Style

        }
    else
        {
            $Style = New-ExcelStyle -HorizontalAlignment Center -NumberFormat '0'
            [PSCustomObject]$Sub |
                ForEach-Object { $_ } |
                Select-Object 'Subscription',
                'Resource Group',
                'Location',
                'Resource Type',
                'Resources Count' | Export-Excel -Path $File -WorksheetName 'Subscriptions' -TableName $TableName -AutoSize -MaxAutoSizeRows 100 -TableStyle $TableStyle -Style $Style
        }

}