<#
.Synopsis
Module for Policy Report

.DESCRIPTION
This script processes and creates the Policy sheet in the Excel report.

.Link
https://github.com/microsoft/ARI/Modules/Private/3.ReportingFunctions/Build-ARIPolicyReport.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Build-ARIPolicyReport {
    param($File ,$Pol, $TableStyle)
    
    # Ensure Pol is an array for safe handling
    if ($null -eq $Pol) {
        $Pol = @()
    } elseif ($Pol -isnot [System.Array]) {
        $Pol = @($Pol)
    }
    
    if ($Pol.Count -gt 0)
        {
            # Memory optimization: Process Policy data in smaller chunks if very large
            $polCount = $Pol.Count
            Write-Debug "Building Policy sheet with $polCount policy record(s)"
            
            # Calculate conditional formatting range based on actual data size (limit to prevent memory issues)
            $maxRows = [math]::Min($polCount + 1, 1000)  # Cap at 1000 rows for conditional formatting
            
            $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize:$false -NumberFormat 0  # Disable AutoSize to save memory

            $condtxt = @()
            if ($maxRows -gt 1) {
                $condtxt += New-ConditionalText -Range "B2:B$maxRows" -ConditionalType GreaterThan 0
                $condtxt += New-ConditionalText -Range "C2:C$maxRows" -ConditionalType GreaterThan 0
                $condtxt += New-ConditionalText -Range "H2:H$maxRows" -ConditionalType GreaterThan 0
            }

            # Export Policy data (without AutoSize to reduce memory usage)
            [PSCustomObject]$Pol |
            ForEach-Object { $_ } |
            Select-Object 'Initiative',
            'Initiative Non Compliance Resources',
            'Initiative Non Compliance Policies',
            'Policy',
            'Policy Type',
            'Effect',
            'Compliance Resources',
            'Non Compliance Resources',
            'Unknown Resources',
            'Exempt Resources',
            'Policy Mode',
            'Policy Version',
            'Policy Deprecated',
            'Policy Category' | Export-Excel -Path $File -WorksheetName 'Policy' -AutoSize:$false -TableName 'AzurePolicy' -MoveToStart -ConditionalText $condtxt -TableStyle $tableStyle -Style $Style
            
            Write-Debug "Policy sheet generated successfully with $polCount record(s)"
        }
}