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
    # Get the Worksheets collection directly from EPPlus
    $WorksheetsCollection = $Excel.Workbook.Worksheets

    # Safely filter worksheets - ensure Name property exists
    $Order = $WorksheetsCollection | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -notin 'Overview','Policy', 'Advisor', 'Security Center', 'Subscriptions', 'Quota Usage', 'AdvisorScore', 'Outages', 'Support Tickets', 'Reservation Advisor' } | Select-Object -Property Index, name, @{N = "Dimension"; E = { if ($null -ne $_.dimension) { $_.dimension.Rows - 1 } else { 0 } } } | Sort-Object -Property Dimension -Descending

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
                $targetSheet = $WorksheetsCollection[$Ord.Name]
                if ($null -ne $targetSheet) {
                    if ($Loop -ne 0 -and $Order0.Count -gt ($Loop - 1) -and $null -ne $Order0[$Loop - 1] -and $null -ne $Order0[$Loop - 1].Name) {
                        try {
                            $afterSheet = $WorksheetsCollection[$Order0[$Loop - 1].Name]
                            if ($null -ne $afterSheet) {
                                # Use EPPlus Position property to move sheet after target
                                $targetSheet.Position = $afterSheet.Position + 1
                            }
                        } catch {
                            Write-Debug "  Warning: Could not move sheet $($Ord.Name): $_"
                        }
                    }
                    if ($Loop -eq 0 -and $null -ne $Order[0] -and $null -ne $Order[0].Name) {
                        try {
                            $afterSheet = $WorksheetsCollection[$Order[0].Name]
                            if ($null -ne $afterSheet) {
                                # Use EPPlus Position property to move sheet after target
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
    # Reorder special sheets after Overview using EPPlus Position property
    $overviewSheet = $WorksheetsCollection['Overview']
    if ($null -ne $overviewSheet) {
        # First, ensure Overview is at position 0
        try {
            if ($overviewSheet.PSObject.Properties.Name -contains 'Position') {
                $overviewSheet.Position = 0
            }
        } catch {
            Write-Debug "  Warning: Could not set Overview position: $_"
        }
        
        # Define sheets to move after Overview in order
        $sheetsToMove = @('Advisor', 'Policy', 'Security Center', 'Quota Usage', 'AdvisorScore', 'Support Tickets', 'Reservation Advisor', 'Subscriptions')
        
        # Collect all sheets that exist
        $sheetsToReorder = @()
        foreach ($sheetName in $sheetsToMove) {
            $sheet = $WorksheetsCollection[$sheetName]
            if ($null -ne $sheet) {
                $sheetsToReorder += [PSCustomObject]@{
                    Name = $sheetName
                    Sheet = $sheet
                }
            }
        }
        
        # Set positions sequentially after Overview (position 0)
        if ($sheetsToReorder.Count -gt 0) {
            $currentPosition = 1  # Start after Overview (position 0)
            
            foreach ($sheetInfo in $sheetsToReorder) {
                try {
                    if ($sheetInfo.Sheet.PSObject.Properties.Name -contains 'Position') {
                        $sheetInfo.Sheet.Position = $currentPosition
                        $currentPosition++
                    } else {
                        # Fallback: Try using MoveAfter if available
                        try {
                            if ($currentPosition -eq 1) {
                                $WorksheetsCollection.MoveAfter($sheetInfo.Name, 'Overview')
                            } else {
                                $prevSheetName = $sheetsToReorder[$currentPosition - 2].Name
                                $WorksheetsCollection.MoveAfter($sheetInfo.Name, $prevSheetName)
                            }
                            $currentPosition++
                        } catch {
                            Write-Debug "  Warning: Could not move $($sheetInfo.Name) sheet (Position property not available)"
                        }
                    }
                } catch {
                    Write-Debug "  Warning: Could not move $($sheetInfo.Name) sheet: $_"
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