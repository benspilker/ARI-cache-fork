<#
.Synopsis
Module for Excel Sheet Ordering

.DESCRIPTION
This script organizes the order of sheets in the Excel report.

.Link
https://github.com/microsoft/ARI/Modules/Private/3.ReportingFunctions/StyleFunctions/Start-ARIExcelOrdening.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>

function Start-ARIExcelOrdening {
    Param($File)

    $Excel = Open-ExcelPackage -Path $File
    $Worksheets = $Excel.Workbook.Worksheets

    # Safely filter worksheets - ensure Name property exists
    $Order = $Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -notin 'Overview','Policy', 'Advisor', 'Security Center', 'Subscriptions', 'Quota Usage', 'AdvisorScore', 'Outages', 'Support Tickets', 'Reservation Advisor' } | Select-Object -Property Index, name, @{N = "Dimension"; E = { if ($null -ne $_.dimension) { $_.dimension.Rows - 1 } else { 0 } } } | Sort-Object -Property Dimension -Descending

    # Safely access Order array elements
    if ($Order -and $Order.Count -gt 0) {
        $firstOrderName = if ($null -ne $Order[0] -and $null -ne $Order[0].Name) { $Order[0].Name } else { $null }
        $lastOrder = $Order | Select-Object -Last 1
        $lastOrderName = if ($null -ne $lastOrder -and $null -ne $lastOrder.Name) { $lastOrder.Name } else { $null }
        
        $Order0 = $Order | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -ne $firstOrderName -and $_.Name -ne $lastOrderName }

        #$Worksheets.MoveAfter(($Order | select-object -Last 1).Name, 'Subscriptions')

        $Loop = 0

        Foreach ($Ord in $Order0) {
            if ($null -ne $Ord -and $null -ne $Ord.Index -and $null -ne $Ord.Name) {
                if ($Loop -ne 0 -and $null -ne $Order0[$Loop - 1] -and $null -ne $Order0[$Loop - 1].Name) {
                    try {
                        $Worksheets.MoveAfter($Ord.Name, $Order0[$Loop - 1].Name)
                    } catch {
                        Write-Debug "  Warning: Could not move sheet $($Ord.Name): $_"
                    }
                }
                if ($Loop -eq 0 -and $null -ne $Order[0] -and $null -ne $Order[0].Name) {
                    try {
                        $Worksheets.MoveAfter($Ord.Name, $Order[0].Name)
                    } catch {
                        Write-Debug "  Warning: Could not move sheet $($Ord.Name): $_"
                    }
                }
            }
            $Loop++
        }
    }

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Validating if Advisor and Policies are included.')
    # Safely check for worksheets with Name property
    if ($null -ne $Worksheets) {
        $advisorSheet = $Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Advisor'} | Select-Object -First 1
        if ($advisorSheet) {
            try {
                $Worksheets.MoveAfter('Advisor', 'Overview')
            } catch {
                Write-Debug "  Warning: Could not move Advisor sheet: $_"
            }
        }
        
        $policySheet = $Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Policy'} | Select-Object -First 1
        if ($policySheet) {
            try {
                $Worksheets.MoveAfter('Policy', 'Overview')
            } catch {
                Write-Debug "  Warning: Could not move Policy sheet: $_"
            }
        }
        
        $securitySheet = $Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Security Center'} | Select-Object -First 1
        if ($securitySheet) {
            try {
                $Worksheets.MoveAfter('Security Center', 'Overview')
            } catch {
                Write-Debug "  Warning: Could not move Security Center sheet: $_"
            }
        }
        
        $quotaSheet = $Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Quota Usage'} | Select-Object -First 1
        if ($quotaSheet) {
            try {
                $Worksheets.MoveAfter('Quota Usage', 'Overview')
            } catch {
                Write-Debug "  Warning: Could not move Quota Usage sheet: $_"
            }
        }
        
        $advisorScoreSheet = $Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'AdvisorScore'} | Select-Object -First 1
        if ($advisorScoreSheet) {
            try {
                $Worksheets.MoveAfter('AdvisorScore', 'Overview')
            } catch {
                Write-Debug "  Warning: Could not move AdvisorScore sheet: $_"
            }
        }
        
        $supportTicketsSheet = $Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Support Tickets'} | Select-Object -First 1
        if ($supportTicketsSheet) {
            try {
                $Worksheets.MoveAfter('Support Tickets', 'Overview')
            } catch {
                Write-Debug "  Warning: Could not move Support Tickets sheet: $_"
            }
        }
        
        $reservationSheet = $Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Reservation Advisor'} | Select-Object -First 1
        if ($reservationSheet) {
            try {
                $Worksheets.MoveAfter('Reservation Advisor', 'Overview')
            } catch {
                Write-Debug "  Warning: Could not move Reservation Advisor sheet: $_"
            }
        }
        
        $subscriptionsSheet = $Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Subscriptions'} | Select-Object -First 1
        if ($subscriptionsSheet) {
            try {
                $Worksheets.MoveAfter('Subscriptions','Overview')
            } catch {
                Write-Debug "  Warning: Could not move Subscriptions sheet: $_"
            }
        }
    }

    $WS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Overview' }

    $WS.SetValue(75,70,'')
    $WS.SetValue(76,70,'')
    $WS.View.ShowGridLines = $false

    $TabDraw = $WS.Drawings.AddShape('TP00', 'RoundRect')
    $TabDraw.SetSize(130 , 78)
    $TabDraw.SetPosition(1, 0, 0, 0)
    $TabDraw.TextAlignment = 'Center'

    Close-ExcelPackage $Excel

}