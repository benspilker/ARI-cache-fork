<#
.Synopsis
Module for Excel Chart Creation

.DESCRIPTION
This script creates charts in the Overview sheet of the Excel report.

.Link
https://github.com/microsoft/ARI/Modules/Private/3.ReportingFunctions/StyleFunctions/Build-ARIExcelChart.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola
#>
function Build-ARIExcelChart {
    Param($Excel, $Overview, $IncludeCosts)
    
    # Helper function to get source range for pivot table
    # EPPlus requires explicit range reference in pivot cache definitions
    # CRITICAL: EPPlus may ignore SourceRange when SourceWorkSheet is also provided
    # So we need to format SourceRange with sheet name: 'SheetName!A1:Z100'
    function Get-PivotTableSourceRange {
        param($Worksheet)
        if ($null -eq $Worksheet) { 
            Write-Debug "  [Get-PivotTableSourceRange] Worksheet is null"
            return $null 
        }
        
        $sheetName = $Worksheet.Name
        $range = $null
        
        # Prefer table address if worksheet has a table
        if ($Worksheet.Tables.Count -gt 0 -and $null -ne $Worksheet.Tables[0].Address) {
            $range = $Worksheet.Tables[0].Address.Address
            Write-Debug "  [Get-PivotTableSourceRange] Using table address: $range"
        }
        # Fallback to worksheet dimension
        elseif ($null -ne $Worksheet.Dimension) {
            $range = $Worksheet.Dimension.Address
            Write-Debug "  [Get-PivotTableSourceRange] Using worksheet dimension: $range"
        }
        
        if ($null -ne $range) {
            # CRITICAL: EPPlus may ignore SourceRange when SourceWorkSheet is also provided
            # We need to format SourceRange with sheet name so EPPlus can identify the source
            # Format: 'SheetName'!A1:Z100 (Excel formula format)
            $formattedRange = "'$sheetName'!$range"
            Write-Debug "  [Get-PivotTableSourceRange] Formatted range with sheet name: $formattedRange"
            return $formattedRange
        }
        
        Write-Debug "  [Get-PivotTableSourceRange] No table or dimension found for worksheet: $sheetName"
        return $null
    }

    $WS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Overview' } | Select-Object -First 1
    if ($null -eq $WS) {
        Write-Error "Overview worksheet not found in Excel workbook"
        return
    }

    $DrawP00 = $WS.Drawings | Where-Object { $_.Name -eq 'TP00' } | Select-Object -First 1
    if ($null -ne $DrawP00) {
        $P00Name = 'Reported Resources'
        $DrawP00.RichText.Add($P00Name).Size = 16
    } else {
        Write-Debug "  Warning: Drawing 'TP00' not found on Overview sheet - skipping"
    }

    if($IncludeCosts.IsPresent)
        {
            # Safely check if Subscriptions worksheet exists before accessing
            $SubscriptionsWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Subscriptions' } | Select-Object -First 1
            if ($null -ne $SubscriptionsWS) {
                # CRITICAL: Get source range for pivot table
                # EPPlus requires explicit range reference to create valid pivot cache XML
                $sourceRange = Get-PivotTableSourceRange -Worksheet $SubscriptionsWS
                
                $PTParams = @{
                    PivotTableName          = "P00"
                    Address                 = $excel.Overview.cells["BA5"] # top-left corner of the table
                    SourceWorkSheet         = $SubscriptionsWS
                    PivotRows               = @("Subscription")
                    PivotData               = @{"Cost" = "Sum" }
                    PivotColumns            = @("Month")
                    PivotTableStyle         = $TableStyle
                    IncludePivotChart       = $true
                    ChartType               = "ColumnStacked3D"
                    ChartRow                = 1 # place the chart below row 22nd
                    ChartColumn             = 9
                    Activate                = $true
                    PivotNumberFormat       = 'Currency'
                    PivotFilter             = 'Resource Type'
                    PivotTotals             = 'Rows'
                    ShowCategory            = $false
                    NoLegend                = $true
                    ChartTitle              = 'Azure Cost per Subscription'
                    ShowPercent             = $true
                    ChartHeight             = 400
                    ChartWidth              = 950
                    ChartRowOffSetPixels    = 0
                    ChartColumnOffSetPixels = 5
                }
                # Add SourceRange if we found a valid range (prevents empty 'ref' attribute in pivot cache)
                if ($null -ne $sourceRange) {
                    $PTParams['SourceRange'] = $sourceRange
                }
                Add-PivotTable @PTParams
            } else {
                Write-Debug "  Warning: Subscriptions worksheet not found - skipping P00 Cost per Subscription PivotTable"
            }
        }
    elseif ($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Reservation Advisor' }) {
        $ReservationAdvisorWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Reservation Advisor' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $ReservationAdvisorWS
        
        $PTParams = @{
            PivotTableName          = "P00"
            Address                 = $excel.Overview.cells["BA5"] # top-left corner of the table
            SourceWorkSheet         = $ReservationAdvisorWS
            PivotRows               = @("Subscription")
            PivotData               = @{"Net Savings" = "Sum" }
            PivotTableStyle         = $TableStyle
            IncludePivotChart       = $true
            ChartType               = "ColumnStacked3D"
            ChartRow                = 1 # place the chart below row 22nd
            ChartColumn             = 9
            Activate                = $true
            PivotNumberFormat       = '$#'
            PivotFilter             = 'Recommended Size'
            PivotTotals             = 'Both'
            ShowCategory            = $false
            NoLegend                = $true
            ChartTitle              = 'Potential Net Savings (VM Reservation)'
            ShowPercent             = $true
            ChartHeight             = 400
            ChartWidth              = 950
            ChartRowOffSetPixels    = 0
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams
    }
    else
        {
            Add-ExcelChart -Worksheet $excel.Overview -ChartType Area3D -XRange "AzureTabs[Name]" -YRange "AzureTabs[Size]" -SeriesHeader 'Resources', 'Count' -Column 9 -Row 1 -Height 400 -Width 950 -RowOffSetPixels 0 -ColumnOffSetPixels 5 -NoLegend
        }

    if($IncludeCosts.IsPresent)
        {
            # Safely check if Subscriptions worksheet exists before accessing
            $SubscriptionsWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Subscriptions' } | Select-Object -First 1
            if ($null -ne $SubscriptionsWS) {
                $sourceRange = Get-PivotTableSourceRange -Worksheet $SubscriptionsWS
                $P0Name = 'CostPerRegion'
                $PTParams = @{
                    PivotTableName          = "P0"
                    Address                 = $excel.Overview.cells["BG5"] # top-left corner of the table
                    SourceWorkSheet         = $SubscriptionsWS
                    PivotRows               = @("Location")
                    PivotData               = @{"Cost" = "Sum" }
                    PivotColumns            = @("Month")
                    PivotTableStyle         = $tableStyle
                    IncludePivotChart       = $true
                    ChartType               = "BarStacked3D"
                    ChartRow                = 13 # place the chart below row 22nd
                    ChartColumn             = 2
                    Activate                = $true
                    PivotNumberFormat       = 'Currency'
                    PivotFilter             = 'Subscription'
                    ChartTitle              = 'Cost by Azure Region'
                    ShowPercent             = $true
                    ChartHeight             = 275
                    ChartWidth              = 445
                    ChartRowOffSetPixels    = 5
                    ChartColumnOffSetPixels = 5
                }
                if ($null -ne $sourceRange) {
                    $PTParams['SourceRange'] = $sourceRange
                }
                Add-PivotTable @PTParams -NoLegend
            } else {
                Write-Debug "  Warning: Subscriptions worksheet not found - skipping P0 CostPerRegion PivotTable"
                $P0Name = $null
            }
    }
    elseif (($Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Outages' }) -and $Overview -eq 1) {
        # Safely check if Outages worksheet exists before accessing
        $OutagesWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Outages' } | Select-Object -First 1
        if ($null -ne $OutagesWS) {
            $sourceRange = Get-PivotTableSourceRange -Worksheet $OutagesWS
            $P0Name = 'Outages'
            $PTParams = @{
                PivotTableName          = "P0"
                Address                 = $excel.Overview.cells["BG5"] # top-left corner of the table
                SourceWorkSheet         = $OutagesWS
                PivotRows               = @("Event Type")
                PivotData               = @{"Outage ID" = "Count" }
                PivotTableStyle         = $tableStyle
                IncludePivotChart       = $true
                ChartType               = "BarStacked3D"
                ChartRow                = 13 # place the chart below row 22nd
                ChartColumn             = 2
                Activate                = $true
                PivotFilter             = 'Subscription', 'Status'
                ChartTitle              = 'Outages (Last 6 Months)'
                ShowPercent             = $true
                ChartHeight             = 275
                ChartWidth              = 445
                ChartRowOffSetPixels    = 5
                ChartColumnOffSetPixels = 5
            }
            if ($null -ne $sourceRange) {
                $PTParams['SourceRange'] = $sourceRange
            }
            Add-PivotTable @PTParams -NoLegend
        } else {
            Write-Debug "  Warning: Outages worksheet not found - skipping P0 Outages PivotTable"
            $P0Name = $null
        }
    }
    elseif (($Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Advisor' }) -and $Overview -eq 2) {
        # Safely check if Advisor worksheet exists before accessing
        $AdvisorWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Advisor' } | Select-Object -First 1
        if ($null -ne $AdvisorWS) {
            $sourceRange = Get-PivotTableSourceRange -Worksheet $AdvisorWS
            $P0Name = 'Advisories'
            $PTParams = @{
                PivotTableName          = "P0"
                Address                 = $excel.Overview.cells["BG5"] # top-left corner of the table
                SourceWorkSheet         = $AdvisorWS
                PivotRows               = @("Category")
                PivotData               = @{"Category" = "Count" }
                PivotTableStyle         = $tableStyle
                IncludePivotChart       = $true
                ChartType               = "BarStacked3D"
                ChartRow                = 13 # place the chart below row 22nd
                ChartColumn             = 2
                Activate                = $true
                PivotFilter             = 'Impact'
                ChartTitle              = 'Advisor'
                ShowPercent             = $true
                ChartHeight             = 275
                ChartWidth              = 445
                ChartRowOffSetPixels    = 5
                ChartColumnOffSetPixels = 5
            }
            if ($null -ne $sourceRange) {
                $PTParams['SourceRange'] = $sourceRange
            }
            Add-PivotTable @PTParams -NoLegend
        } else {
            Write-Debug "  Warning: Advisor worksheet not found - skipping P0 Advisories PivotTable"
            $P0Name = $null
        }
    }
    else {
        # Safely check if Public IPs worksheet exists before accessing
        $PublicIPsWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Public IPs' } | Select-Object -First 1
        if ($null -ne $PublicIPsWS) {
            $sourceRange = Get-PivotTableSourceRange -Worksheet $PublicIPsWS
            $P0Name = 'Public IPs'
            $PTParams = @{
                PivotTableName          = "P0"
                Address                 = $excel.Overview.cells["BG5"] # top-left corner of the table
                SourceWorkSheet         = $PublicIPsWS
                PivotRows               = @("Use")
                PivotData               = @{"Use" = "Count" }
                PivotTableStyle         = $tableStyle
                IncludePivotChart       = $true
                ChartType               = "BarStacked3D"
                ChartRow                = 13 # place the chart below row 22nd
                ChartColumn             = 2
                Activate                = $true
                PivotFilter             = 'location'
                ChartTitle              = 'Public IPs'
                ShowPercent             = $true
                ChartHeight             = 275
                ChartWidth              = 445
                ChartRowOffSetPixels    = 5
                ChartColumnOffSetPixels = 5
            }
            if ($null -ne $sourceRange) {
                $PTParams['SourceRange'] = $sourceRange
            }
            Add-PivotTable @PTParams -NoLegend
        } else {
            Write-Debug "  Warning: Public IPs worksheet not found - skipping P0 Public IPs PivotTable"
            $P0Name = $null
        }
    }

    $DrawP0 = $WS.Drawings | Where-Object { $_.Name -eq 'TP0' }
    if ($null -ne $P0Name) {
        $DrawP0.RichText.Add($P0Name) | Out-Null
    } else {
        Write-Debug "  Warning: P0Name not set - skipping TP0 rich text"
    }

    if($IncludeCosts.IsPresent)
        {
            # Safely check if Subscriptions worksheet exists before accessing
            $SubscriptionsWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Subscriptions' } | Select-Object -First 1
            if ($null -ne $SubscriptionsWS) {
                $sourceRange = Get-PivotTableSourceRange -Worksheet $SubscriptionsWS
                $P1Name = 'TotalCostsPerSubscription'
                $PTParams = @{
                    PivotTableName          = "P1"
                    Address                 = $excel.Overview.cells["DK6"] # top-left corner of the table
                    SourceWorkSheet         = $SubscriptionsWS
                    PivotRows               = @('Month')
                    PivotColumns            = @("Resource Type")
                    PivotData               = @{"Cost" = "Sum" }
                    PivotTableStyle         = $tableStyle
                    IncludePivotChart       = $true
                ChartType               = "BarClustered"
                ChartRow                = 27 # place the chart below row 22nd
                ChartColumn             = 2
                Activate                = $true
                PivotNumberFormat       = 'Currency'
                ShowCategory            = $false
                ChartTitle              = 'Cost by Resource Type'
                NoLegend                = $true
                ShowPercent             = $true
                ChartHeight             = 655
                ChartWidth              = 570
                ChartRowOffSetPixels    = 5
                ChartColumnOffSetPixels = 5
                }
                if ($null -ne $sourceRange) {
                    $PTParams['SourceRange'] = $sourceRange
                }
                Add-PivotTable @PTParams
            } else {
                Write-Debug "  Warning: Subscriptions worksheet not found - skipping P1 TotalCostsPerSubscription PivotTable"
            }
        }
    elseif (($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'AdvisorScore' }) -and $Overview -eq 1) {
        $AdvisorScoreWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'AdvisorScore' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $AdvisorScoreWS
        $P1Name = 'AdvisorScore'
        $PTParams = @{
            PivotTableName          = "P1"
            Address                 = $excel.Overview.cells["DK6"] # top-left corner of the table
            SourceWorkSheet         = $AdvisorScoreWS
            PivotRows               = @("Category")
            PivotData               = @{"Latest Score (%)" = "average" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "BarClustered"
            ChartRow                = 27 # place the chart below row 22nd
            ChartColumn             = 2
            Activate                = $true
            #PivotNumberFormat       = '0'
            ShowCategory            = $false
            PivotFilter             = 'Subscription'
            ChartTitle              = 'Advisor Score (%)'
            NoLegend                = $true
            ShowPercent             = $true
            ChartHeight             = 655
            ChartWidth              = 570
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams
    }
    elseif (($Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Subscriptions' }) -and $Overview -eq 2) {
        # Safely check if Subscriptions worksheet exists before accessing
        $SubscriptionsWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Subscriptions' } | Select-Object -First 1
        if ($null -ne $SubscriptionsWS) {
            $sourceRange = Get-PivotTableSourceRange -Worksheet $SubscriptionsWS
            $P1Name = 'Subscriptions'
            $PTParams = @{
                PivotTableName          = "P1"
                Address                 = $excel.Overview.cells["DK6"] # top-left corner of the table
                SourceWorkSheet         = $SubscriptionsWS
                PivotRows               = @("Subscription")
                PivotData               = @{"Resources Count" = "sum" }
                PivotTableStyle         = $tableStyle
                IncludePivotChart       = $true
                ChartType               = "BarClustered"
            ChartRow                = 27 # place the chart below row 22nd
            ChartColumn             = 2
            Activate                = $true
            PivotFilter             = 'Resource Group'
            ChartTitle              = 'Resources by Subscription'
            NoLegend                = $true
            ShowPercent             = $true
            ChartHeight             = 655
            ChartWidth              = 570
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
            }
            if ($null -ne $sourceRange) {
                $PTParams['SourceRange'] = $sourceRange
            }
            Add-PivotTable @PTParams
        } else {
            Write-Debug "  Warning: Subscriptions worksheet not found - skipping P1 PivotTable"
        }
    }
    elseif (($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Quota Usage' }) -and $Overview -eq 1) {
        $QuotaUsageWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Quota Usage' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $QuotaUsageWS
        $P1Name = 'Quota Usage'
        $PTParams = @{
            PivotTableName          = "P1"
            Address                 = $excel.Overview.cells["DK6"] # top-left corner of the table
            SourceWorkSheet         = $QuotaUsageWS
            PivotRows               = @("Region")
            PivotData               = @{"vCPUs Available" = "Sum" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "BarClustered"
            ChartRow                = 27 # place the chart below row 22nd
            ChartColumn             = 2
            Activate                = $true
            PivotFilter             = 'Limit'
            ChartTitle              = 'Available Quota (vCPUs)'
            NoLegend                = $true
            ShowPercent             = $true
            ChartHeight             = 655
            ChartWidth              = 570
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams
    }
    else {
        $VirtualNetworksWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Virtual Networks' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $VirtualNetworksWS
        $P1Name = 'Virtual Networks'
        $PTParams = @{
            PivotTableName          = "P1"
            Address                 = $excel.Overview.cells["DK6"] # top-left corner of the table
            SourceWorkSheet         = $VirtualNetworksWS
            PivotRows               = @("Name")
            PivotData               = @{"Available IPs" = "Sum" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "BarClustered"
            ChartRow                = 27 # place the chart below row 22nd
            ChartColumn             = 2
            Activate                = $true
            PivotFilter             = 'Location'
            ChartTitle              = 'Available IPs (Per Virtual Network)'
            NoLegend                = $true
            ShowPercent             = $true
            ChartHeight             = 655
            ChartWidth              = 570
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams
    }
    $DrawP1 = $WS.Drawings | Where-Object { $_.Name -eq 'TP1' }
    $DrawP1.RichText.Add($P1Name) | Out-Null

    if (($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Policy' }) -and $Overview -eq 1) {
        $PolicyWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Policy' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $PolicyWS
        $P2Name = 'Policy'
        $PTParams = @{
            PivotTableName          = "P2"
            Address                 = $excel.Overview.cells["BT5"] # top-left corner of the table
            SourceWorkSheet         = $PolicyWS
            PivotRows               = @("Policy Category")
            PivotData               = @{"Policy" = "Count" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "ColumnStacked3D"
            ChartRow                = 21 # place the chart below row 22nd
            ChartColumn             = 11
            Activate                = $true
            PivotFilter             = 'Policy Type'
            ChartTitle              = 'Policies by Category'
            ShowPercent             = $true
            ChartHeight             = 255
            ChartWidth              = 315
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams -NoLegend
    }
    elseif (($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Advisor' }) -and $Overview -eq 2) {
        $AdvisorWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Advisor' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $AdvisorWS
        $P2Name = 'Annual Savings'
        $PTParams = @{
            PivotTableName          = "P2"
            Address                 = $excel.Overview.cells["BT5"] # top-left corner of the table
            SourceWorkSheet         = $AdvisorWS
            PivotRows               = @("Savings Currency")
            PivotData               = @{"Annual Savings" = "Sum" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "ColumnStacked3D"
            ChartRow                = 21 # place the chart below row 22nd
            ChartColumn             = 11
            Activate                = $true
            ChartTitle              = 'Potential Savings'
            PivotFilter             = 'Savings Region'
            ShowPercent             = $true
            ChartHeight             = 255
            ChartWidth              = 315
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
            PivotNumberFormat       = '#,##0.00'
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams -NoLegend
    }
    else {
        $VirtualNetworksWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Virtual Networks' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $VirtualNetworksWS
        $P2Name = 'Virtual Networks'
        $PTParams = @{
            PivotTableName          = "P2"
            Address                 = $excel.Overview.cells["BT5"] # top-left corner of the table
            SourceWorkSheet         = $VirtualNetworksWS
            PivotRows               = @("Location")
            PivotData               = @{"Location" = "Count" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "ColumnStacked3D"
            ChartRow                = 21 # place the chart below row 22nd
            ChartColumn             = 11
            Activate                = $true
            ChartTitle              = 'Virtual Networks'
            PivotFilter             = 'Subscription'
            ShowPercent             = $true
            ChartHeight             = 255
            ChartWidth              = 315
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams -NoLegend
    }

    $DrawP2 = $WS.Drawings | Where-Object { $_.Name -eq 'TP2' }
    $DrawP2.RichText.Add($P2Name) | Out-Null

    if (($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Support Tickets' }) -and $Overview -eq 1) {
        $SupportTicketsWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Support Tickets' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $SupportTicketsWS
        $P3Name = 'Support Tickets'
        $PTParams = @{
            PivotTableName          = "P3"
            Address                 = $excel.Overview.cells["BZ5"] # top-left corner of the table
            SourceWorkSheet         = $SupportTicketsWS
            PivotRows               = @("Status")
            PivotData               = @{"Status" = "Count" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "Pie3D"
            ChartRow                = 34 # place the chart below row 22nd
            ChartColumn             = 11
            Activate                = $true
            PivotFilter             = 'Current Severity'
            ChartTitle              = 'Support Tickets'
            ShowPercent             = $true
            ChartHeight             = 255
            ChartWidth              = 315
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams
    }
    elseif (($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'AKS' }) -and $Overview -eq 2) {
        $AKSWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'AKS' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $AKSWS
        $P3Name = 'Azure Kubernetes'
        $PTParams = @{
            PivotTableName          = "P3"
            Address                 = $excel.Overview.cells["BZ5"] # top-left corner of the table
            SourceWorkSheet         = $AKSWS
            PivotRows               = @("Kubernetes Version")
            PivotData               = @{"Clusters" = "Count" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "Pie3D"
            ChartRow                = 34 # place the chart below row 22nd
            ChartColumn             = 11
            Activate                = $true
            ChartTitle              = 'AKS Versions'
            PivotFilter             = 'Node Pool Size'
            ShowPercent             = $true
            ChartHeight             = 255
            ChartWidth              = 315
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams
    }
    else {
        $StorageAccountsWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Storage Accounts' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $StorageAccountsWS
        $P3Name = 'Storage Accounts'
        $PTParams = @{
            PivotTableName          = "P3"
            Address                 = $excel.Overview.cells["BZ5"] # top-left corner of the table
            SourceWorkSheet         = $StorageAccountsWS
            PivotRows               = @("Tier")
            PivotData               = @{"Tier" = "Count" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "Pie3D"
            ChartRow                = 34 # place the chart below row 22nd
            ChartColumn             = 11
            Activate                = $true
            PivotFilter             = 'SKU'
            ChartTitle              = 'Storage Accounts'
            ShowPercent             = $true
            ChartHeight             = 255
            ChartWidth              = 315
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams
    }
    $DrawP3 = $WS.Drawings | Where-Object { $_.Name -eq 'TP3' }
    $DrawP3.RichText.Add($P3Name) | Out-Null

    # Removed duplicate Outages chart (P4) - P0 chart already shows Outages data
    # Keeping P0 chart ('Outages (Last 6 Months)') as it provides better insights
    # Initialize P4Name to null, then check for other P4 chart options
    $P4Name = $null
    if (($Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Quota Usage' }) -and $Overview -eq 2) {
        # Safely check if Quota Usage worksheet exists before accessing
        $QuotaUsageWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Quota Usage' } | Select-Object -First 1
        if ($null -ne $QuotaUsageWS) {
            $sourceRange = Get-PivotTableSourceRange -Worksheet $QuotaUsageWS
            $P4Name = 'Quota Usage'
            $PTParams = @{
                PivotTableName          = "P4"
                Address                 = $excel.Overview.cells["CF5"] # top-left corner of the table
                SourceWorkSheet         = $QuotaUsageWS
                PivotRows               = @("Region")
                PivotData               = @{"vCPUs Available" = "Sum" }
                PivotTableStyle         = $tableStyle
                IncludePivotChart       = $true
                ChartType               = "ColumnStacked3D"
                ChartRow                = 47 # place the chart below row 22nd
                ChartColumn             = 11
                Activate                = $true
                PivotFilter             = 'Limit'
                ChartTitle              = 'Available Quota (vCPUs)'
                ShowPercent             = $true
                ChartHeight             = 255
                ChartWidth              = 315
                ChartRowOffSetPixels    = 5
                ChartColumnOffSetPixels = 5
            }
            if ($null -ne $sourceRange) {
                $PTParams['SourceRange'] = $sourceRange
            }
            Add-PivotTable @PTParams -NoLegend
        } else {
            Write-Debug "  Warning: Quota Usage worksheet not found - skipping P4 Quota Usage PivotTable"
            $P4Name = $null
        }
    }
    else {
        # Safely check if Disks worksheet exists before accessing
        $DisksWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Disks' } | Select-Object -First 1
        if ($null -ne $DisksWS) {
            $sourceRange = Get-PivotTableSourceRange -Worksheet $DisksWS
            $P4Name = 'VM Disks'
            $PTParams = @{
                PivotTableName          = "P4"
                Address                 = $excel.Overview.cells["CF5"] # top-left corner of the table
                SourceWorkSheet         = $DisksWS
                PivotRows               = @("Disk State")
                PivotData               = @{"Disk State" = "Count" }
                PivotTableStyle         = $tableStyle
                IncludePivotChart       = $true
                ChartType               = "ColumnStacked3D"
                ChartRow                = 47 # place the chart below row 22nd
                ChartColumn             = 11
                Activate                = $true
                PivotFilter             = 'SKU'
                ChartTitle              = 'VM Disks'
                ShowPercent             = $true
                ChartHeight             = 255
                ChartWidth              = 315
                ChartRowOffSetPixels    = 5
                ChartColumnOffSetPixels = 5
            }
            if ($null -ne $sourceRange) {
                $PTParams['SourceRange'] = $sourceRange
            }
            Add-PivotTable @PTParams -NoLegend
        } else {
            Write-Debug "  Warning: Disks worksheet not found - skipping P4 VM Disks PivotTable"
            $P4Name = $null
        }
    }

    $DrawP4 = $WS.Drawings | Where-Object { $_.Name -eq 'TP4' }
    if ($null -ne $P4Name) {
        $DrawP4.RichText.Add($P4Name) | Out-Null
    } else {
        Write-Debug "  Warning: P4Name not set - skipping TP4 rich text"
    }

    if ($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Virtual Machines' }) {
        $VirtualMachinesWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Virtual Machines' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $VirtualMachinesWS
        $P5Name = 'Virtual Machines'
        $PTParams = @{
            PivotTableName          = "P5"
            Address                 = $excel.Overview.cells["CL7"] # top-left corner of the table
            SourceWorkSheet         = $VirtualMachinesWS
            PivotRows               = @("VM Size")
            PivotData               = @{"Resource U" = "Sum" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "BarClustered"
            ChartRow                = 21 # place the chart below row 22nd
            ChartColumn             = 16
            Activate                = $true
            NoLegend                = $true
            ChartTitle              = 'Virtual Machines by Series'
            PivotFilter             = 'OS Type', 'Location', 'Power State'
            ShowPercent             = $true
            ChartHeight             = 775
            ChartWidth              = 502
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams
    } else {
        $VirtualNetworksWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Virtual Networks' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $VirtualNetworksWS
        $P5Name = 'Virtual Networks'
        $PTParams = @{
            PivotTableName          = "P5"
            Address                 = $excel.Overview.cells["CL7"] # top-left corner of the table
            SourceWorkSheet         = $VirtualNetworksWS
            PivotRows               = @("Name")
            PivotData               = @{"Available IPs" = "Sum" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "BarClustered"
            ChartRow                = 21 # place the chart below row 22nd
            ChartColumn             = 16
            Activate                = $true
            NoLegend                = $true
            ChartTitle              = 'Available IPs (Per Virtual Network)'
            PivotFilter             = 'Location'
            ShowPercent             = $true
            ChartHeight             = 775
            ChartWidth              = 502
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 5
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams
    }

    $DrawP5 = $WS.Drawings | Where-Object { $_.Name -eq 'TP5' }
    $DrawP5.RichText.Add($P5Name) | Out-Null

    # Safely check if Subscriptions worksheet exists before accessing
    $SubscriptionsWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Subscriptions' } | Select-Object -First 1
    
    # Initialize P6Name with a default value
    $P6Name = $null
    
    if($IncludeCosts.IsPresent)
        {
            if ($null -ne $SubscriptionsWS) {
                $sourceRange = Get-PivotTableSourceRange -Worksheet $SubscriptionsWS
                $P6Name = 'Cost per Month'
                $PTParams = @{
                    PivotTableName          = "P6"
                    Address                 = $excel.Overview.cells["CR5"] # top-left corner of the table
                    SourceWorkSheet         = $SubscriptionsWS
                    PivotRows               = @("Month")
                    PivotData               = @{"Cost" = "sum" }
                    PivotTableStyle         = $tableStyle
                    IncludePivotChart       = $true
                    ChartType               = "ColumnStacked3D"
                    ChartRow                = 1 # place the chart below row 22nd
                    ChartColumn             = 24
                    Activate                = $true
                    PivotNumberFormat       = 'Currency'
                    PivotFilter             = 'Subscription'
                    ChartTitle              = 'Cost per Month'
                    NoLegend                = $true
                    ShowPercent             = $true
                    ChartHeight             = 400
                    ChartWidth              = 315
                    ChartRowOffSetPixels    = 0
                    ChartColumnOffSetPixels = 0
                }
                if ($null -ne $sourceRange) {
                    $PTParams['SourceRange'] = $sourceRange
                }
                Add-PivotTable @PTParams
            } else {
                Write-Debug "  Warning: Subscriptions worksheet not found - skipping P6 Cost per Month PivotTable"
            }
        }
    else
        {
            if ($null -ne $SubscriptionsWS) {
                $sourceRange = Get-PivotTableSourceRange -Worksheet $SubscriptionsWS
                $P6Name = 'Resources by Location'
                $PTParams = @{
                    PivotTableName          = "P6"
                    Address                 = $excel.Overview.cells["CR5"] # top-left corner of the table
                    SourceWorkSheet         = $SubscriptionsWS
                    PivotRows               = @("Location")
                    PivotData               = @{"Resources Count" = "sum" }
                    PivotTableStyle         = $tableStyle
                    IncludePivotChart       = $true
                    ChartType               = "ColumnStacked3D"
                    ChartRow                = 1 # place the chart below row 22nd
                    ChartColumn             = 24
                    Activate                = $true
                    PivotFilter             = 'Resource Type'
                    ChartTitle              = 'Resources by Location'
                    NoLegend                = $true
                    ShowPercent             = $true
                    ChartHeight             = 400
                    ChartWidth              = 315
                    ChartRowOffSetPixels    = 0
                    ChartColumnOffSetPixels = 0
                }
                if ($null -ne $sourceRange) {
                    $PTParams['SourceRange'] = $sourceRange
                }
                Add-PivotTable @PTParams
            } else {
                Write-Debug "  Warning: Subscriptions worksheet not found - skipping P6 Resources by Location PivotTable"
            }
        }

    $DrawP6 = $WS.Drawings | Where-Object { $_.Name -eq 'TP6' }
    if ($null -ne $P6Name) {
        $DrawP6.RichText.Add($P6Name) | Out-Null
    } else {
        Write-Debug "  Warning: P6Name not set - skipping TP6 rich text"
    }

    if ($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Virtual Machines' }) {
        $VirtualMachinesWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Virtual Machines' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $VirtualMachinesWS
        $P7Name = 'Virtual Machines'
        $PTParams = @{
            PivotTableName          = "P7"
            Address                 = $excel.Overview.cells["CY5"] # top-left corner of the table
            SourceWorkSheet         = $VirtualMachinesWS
            PivotRows               = @("OS Type")
            PivotData               = @{"Resource U" = "Sum" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "Pie3D"
            ChartRow                = 21 # place the chart below row 22nd
            ChartColumn             = 24
            Activate                = $true
            NoLegend                = $true
            ChartTitle              = 'VMs by OS'
            PivotFilter             = 'Location'
            ShowPercent             = $true
            ChartHeight             = 255
            ChartWidth              = 315
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 0
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams
    }

    $DrawP7 = $WS.Drawings | Where-Object { $_.Name -eq 'TP7' }
    $DrawP7.RichText.Add($P7Name) | Out-Null

    if ($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Advisor' }) {
        $AdvisorWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Advisor' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $AdvisorWS
        $P8Name = 'Advisories'
        $PTParams = @{
            PivotTableName          = "P8"
            Address                 = $excel.Overview.cells["DE5"] # top-left corner of the table
            SourceWorkSheet         = $AdvisorWS
            PivotRows               = @("Impact")
            PivotData               = @{"Impact" = "Count" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "BarStacked3D"
            ChartRow                = 34
            ChartColumn             = 24
            Activate                = $true
            PivotFilter             = 'Category'
            ChartTitle              = 'Advisor'
            ShowPercent             = $true
            ChartHeight             = 255
            ChartWidth              = 315
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 0
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams -NoLegend
    }
    elseif ($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Load Balancers' }) {
        $LoadBalancersWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Load Balancers' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $LoadBalancersWS
        $P8Name = 'Load Balancers'
        $PTParams = @{
            PivotTableName          = "P8"
            Address                 = $excel.Overview.cells["DE5"] # top-left corner of the table
            SourceWorkSheet         = $LoadBalancersWS
            PivotRows               = @("Usage")
            PivotData               = @{"Usage" = "Count" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "BarStacked3D"
            ChartRow                = 34
            ChartColumn             = 24
            Activate                = $true
            PivotFilter             = 'Location'
            ChartTitle              = 'Load Balancers'
            ShowPercent             = $true
            ChartHeight             = 255
            ChartWidth              = 315
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 0
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams -NoLegend
    }
    elseif ($Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Virtual Machines' }) {
        $VirtualMachinesWS = $Excel.Workbook.Worksheets | Where-Object { $_.Name -eq 'Virtual Machines' } | Select-Object -First 1
        $sourceRange = Get-PivotTableSourceRange -Worksheet $VirtualMachinesWS
        $P8Name = 'VMs per Region'
        $PTParams = @{
            PivotTableName          = "P8"
            Address                 = $excel.Overview.cells["DE5"] # top-left corner of the table
            SourceWorkSheet         = $VirtualMachinesWS
            PivotRows               = @("Location")
            PivotData               = @{"Resource U" = "Sum" }
            PivotTableStyle         = $tableStyle
            IncludePivotChart       = $true
            ChartType               = "BarStacked3D"
            ChartRow                = 34
            ChartColumn             = 24
            Activate                = $true
            PivotFilter             = 'Subscription'
            ChartTitle              = 'VMs by Region'
            ShowPercent             = $true
            ChartHeight             = 255
            ChartWidth              = 315
            ChartRowOffSetPixels    = 5
            ChartColumnOffSetPixels = 0
        }
        if ($null -ne $sourceRange) {
            $PTParams['SourceRange'] = $sourceRange
        }
        Add-PivotTable @PTParams -NoLegend
    }
    else{
        # Safely check if Subscriptions worksheet exists before accessing
        $SubscriptionsWS = $Excel.Workbook.Worksheets | Where-Object { $null -ne $_ -and $null -ne $_.Name -and $_.Name -eq 'Subscriptions' } | Select-Object -First 1
        if ($null -ne $SubscriptionsWS) {
            $sourceRange = Get-PivotTableSourceRange -Worksheet $SubscriptionsWS
            $P8Name = 'Resources per Region'
            $PTParams = @{
                PivotTableName          = "P8"
                Address                 = $excel.Overview.cells["DE5"] # top-left corner of the table
                SourceWorkSheet         = $SubscriptionsWS
                PivotRows               = @("Location")
                PivotData               = @{"Resources Count" = "Sum" }
                PivotTableStyle         = $tableStyle
                IncludePivotChart       = $true
                ChartType               = "BarStacked3D"
                ChartRow                = 34
                ChartColumn             = 24
                Activate                = $true
                PivotFilter             = 'Subscription'
                ChartTitle              = 'Resources by Location'
                ShowPercent             = $true
                ChartHeight             = 255
                ChartWidth              = 315
                ChartRowOffSetPixels    = 5
                ChartColumnOffSetPixels = 0
            }
            if ($null -ne $sourceRange) {
                $PTParams['SourceRange'] = $sourceRange
            }
            Add-PivotTable @PTParams -NoLegend
        } else {
            Write-Debug "  Warning: Subscriptions worksheet not found - skipping P8 Resources per Region PivotTable"
            $P8Name = $null
        }
    }

    $DrawP8 = $WS.Drawings | Where-Object { $_.Name -eq 'TP8' }
    $DrawP8.RichText.Add($P8Name) | Out-Null

    # Removed Boot Diagnostics chart (P9) - no data available and provides limited insight
    $P9Name = $null

    $DrawP9 = $WS.Drawings | Where-Object { $_.Name -eq 'TP9' }
    if ($null -ne $DrawP9) {
        if ($null -ne $P9Name) {
            $DrawP9.RichText.Add($P9Name) | Out-Null
        } else {
            Write-Debug "  Warning: P9Name not set - skipping TP9 rich text"
        }
    }

}