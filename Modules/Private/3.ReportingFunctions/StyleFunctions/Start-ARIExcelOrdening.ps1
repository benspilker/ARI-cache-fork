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
                    # Check if Position property exists before using it
                    $hasPositionProperty = $targetSheet.PSObject.Properties.Name -contains 'Position'
                    
                    if ($hasPositionProperty) {
                        if ($Loop -ne 0 -and $Order0.Count -gt ($Loop - 1) -and $null -ne $Order0[$Loop - 1] -and $null -ne $Order0[$Loop - 1].Name) {
                            try {
                                $afterSheet = $WorksheetsCollection[$Order0[$Loop - 1].Name]
                                if ($null -ne $afterSheet -and $afterSheet.PSObject.Properties.Name -contains 'Position') {
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
                                if ($null -ne $afterSheet -and $afterSheet.PSObject.Properties.Name -contains 'Position') {
                                    # Use EPPlus Position property to move sheet after target
                                    $targetSheet.Position = $afterSheet.Position + 1
                                }
                            } catch {
                                Write-Debug "  Warning: Could not move sheet $($Ord.Name): $_"
                            }
                        }
                    } else {
                        # Position property not available - skip ordering for this sheet
                        # This is non-critical, sheets will remain in creation order
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
        # Check if Position property is available
        $positionAvailable = $overviewSheet.PSObject.Properties.Name -contains 'Position'
        
        if ($positionAvailable) {
            # First, ensure Overview is at position 0
            try {
                $overviewSheet.Position = 0
            } catch {
                Write-Debug "  Warning: Could not set Overview position: $_"
            }
            
            # Define sheets to move after Overview in order
            $sheetsToMove = @('Advisor', 'Policy', 'Security Center', 'Quota Usage', 'AdvisorScore', 'Support Tickets', 'Reservation Advisor', 'Subscriptions')
            
            # Collect all sheets that exist and have Position property
            $sheetsToReorder = @()
            foreach ($sheetName in $sheetsToMove) {
                $sheet = $WorksheetsCollection[$sheetName]
                if ($null -ne $sheet -and $sheet.PSObject.Properties.Name -contains 'Position') {
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
                        $sheetInfo.Sheet.Position = $currentPosition
                        $currentPosition++
                    } catch {
                        Write-Debug "  Warning: Could not move $($sheetInfo.Name) sheet: $_"
                    }
                }
            }
        } else {
            # Position property not available - skip ordering (non-critical)
            # Sheets will remain in creation order
            Write-Debug "  Position property not available - sheet ordering skipped (cosmetic only)"
        }
    }

    $WS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Overview' } | Select-Object -First 1
    if ($null -eq $WS) {
        Write-Debug "  Warning: Overview worksheet not found - skipping TP00 drawing creation"
        Close-ExcelPackage $Excel
        return
    }

    # TP00 drawing may have already been created in Start-ARIExcelCustomization before Build-ARIExcelChart
    # Check if it exists before creating it
    $existingTP00 = $WS.Drawings | Where-Object { $_.Name -eq 'TP00' } | Select-Object -First 1
    if ($null -eq $existingTP00) {
        $WS.SetValue(75,70,'')
        $WS.SetValue(76,70,'')
        $WS.View.ShowGridLines = $false

        $TabDraw = $WS.Drawings.AddShape('TP00', 'RoundRect')
        $TabDraw.SetSize(130 , 78)
        $TabDraw.SetPosition(1, 0, 0, 0)
        $TabDraw.TextAlignment = 'Center'
    } else {
        Write-Debug "  TP00 drawing already exists - skipping creation"
    }

    Close-ExcelPackage $Excel

}