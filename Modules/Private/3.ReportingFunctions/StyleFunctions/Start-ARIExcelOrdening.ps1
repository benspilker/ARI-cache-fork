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
    # Safely access Worksheets - ensure it's always an array
    $Worksheets = if ($null -ne $Excel -and $null -ne $Excel.Workbook -and $null -ne $Excel.Workbook.Worksheets) { 
        $Excel.Workbook.Worksheets 
    } else { 
        @() 
    }
    # Ensure Worksheets is an array
    if ($null -eq $Worksheets) {
        $Worksheets = @()
    } elseif ($Worksheets -isnot [System.Array]) {
        $Worksheets = @($Worksheets)
    }

    # Safely filter worksheets - ensure Name property exists
    $Order = $Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -notin 'Overview','Policy', 'Advisor', 'Security Center', 'Subscriptions', 'Quota Usage', 'AdvisorScore', 'Outages', 'Support Tickets', 'Reservation Advisor' } | Select-Object -Property Index, name, @{N = "Dimension"; E = { if ($null -ne $_.dimension) { $_.dimension.Rows - 1 } else { 0 } } } | Sort-Object -Property Dimension -Descending

    # Ensure Order is an array for safe .Count access
    if ($null -eq $Order) {
        $Order = @()
    } elseif ($Order -isnot [System.Array]) {
        $Order = @($Order)
    }

    # Safely access Order array elements
    if ($Order.Count -gt 0) {
        $firstOrderName = if ($null -ne $Order[0] -and $null -ne $Order[0].Name) { $Order[0].Name } else { $null }
        $lastOrder = $Order | Select-Object -Last 1
        $lastOrderName = if ($null -ne $lastOrder -and $null -ne $lastOrder.Name) { $lastOrder.Name } else { $null }
        
        $Order0 = $Order | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -ne $firstOrderName -and $_.Name -ne $lastOrderName }
        
        # Ensure Order0 is an array for safe array access
        if ($null -eq $Order0) {
            $Order0 = @()
        } elseif ($Order0 -isnot [System.Array]) {
            $Order0 = @($Order0)
        }

        #$Worksheets.MoveAfter(($Order | select-object -Last 1).Name, 'Subscriptions')

        $Loop = 0

        Foreach ($Ord in $Order0) {
            if ($null -ne $Ord -and $null -ne $Ord.Index -and $null -ne $Ord.Name) {
                $targetSheet = $Excel.Workbook.Worksheets[$Ord.Name]
                if ($null -ne $targetSheet) {
                    if ($Loop -ne 0 -and $Order0.Count -gt ($Loop - 1) -and $null -ne $Order0[$Loop - 1] -and $null -ne $Order0[$Loop - 1].Name) {
                        try {
                            $afterSheet = $Excel.Workbook.Worksheets[$Order0[$Loop - 1].Name]
                            if ($null -ne $afterSheet) {
                                $targetSheet.Position = $afterSheet.Position + 1
                            }
                        } catch {
                            Write-Debug "  Warning: Could not move sheet $($Ord.Name): $_"
                        }
                    }
                    if ($Loop -eq 0 -and $null -ne $Order[0] -and $null -ne $Order[0].Name) {
                        try {
                            $afterSheet = $Excel.Workbook.Worksheets[$Order[0].Name]
                            if ($null -ne $afterSheet) {
                                $targetSheet.Position = $afterSheet.Position + 1
                            }
                        } catch {
                            Write-Debug "  Warning: Could not move sheet $($Ord.Name): $_"
                        }
                    }
                }
            }
            $Loop++
        }
    }

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Validating if Advisor and Policies are included.')
    # Safely check for worksheets with Name property
    if ($null -ne $Worksheets) {
        $overviewSheet = $Excel.Workbook.Worksheets['Overview']
        if ($null -ne $overviewSheet) {
            $overviewPosition = $overviewSheet.Position
            $currentPosition = $overviewPosition + 1
            
            # Define sheets to move after Overview in order
            $sheetsToMove = @('Advisor', 'Policy', 'Security Center', 'Quota Usage', 'AdvisorScore', 'Support Tickets', 'Reservation Advisor', 'Subscriptions')
            
            foreach ($sheetName in $sheetsToMove) {
                $sheet = $Excel.Workbook.Worksheets[$sheetName]
                if ($null -ne $sheet) {
                    try {
                        $sheet.Position = $currentPosition
                        $currentPosition++
                    } catch {
                        Write-Debug "  Warning: Could not move $sheetName sheet: $_"
                    }
                }
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