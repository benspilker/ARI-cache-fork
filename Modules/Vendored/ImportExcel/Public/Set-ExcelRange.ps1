function Set-ExcelRange {
    [CmdletBinding()]
    [Alias("Set-Format")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system state')]
    param(
        [Parameter(ValueFromPipeline = $true,Position=0)]
        [Alias("Address")]
        $Range ,
        [OfficeOpenXml.ExcelWorksheet]$Worksheet ,
        [Alias("NFormat")]
        $NumberFormat,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderAround,
        $BorderColor=[System.Drawing.Color]::Black,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderBottom,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderTop,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderLeft,
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderRight,
        [Alias('ForegroundColor')]
        $FontColor,
        $Value,
        $Formula,
        [Switch]$ArrayFormula,
        [Switch]$ResetFont,
        [Switch]$Bold,
        [Switch]$Italic,
        [Switch]$Underline,
        [OfficeOpenXml.Style.ExcelUnderLineType]$UnderLineType = [OfficeOpenXml.Style.ExcelUnderLineType]::Single,
        [Switch]$StrikeThru,
        [OfficeOpenXml.Style.ExcelVerticalAlignmentFont]$FontShift,
        [String]$FontName,
        [float]$FontSize,
        $BackgroundColor,
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern = [OfficeOpenXml.Style.ExcelFillStyle]::Solid ,
        [Alias("PatternColour")]
        $PatternColor,
        [Switch]$WrapText,
        [OfficeOpenXml.Style.ExcelHorizontalAlignment]$HorizontalAlignment,
        [OfficeOpenXml.Style.ExcelVerticalAlignment]$VerticalAlignment,
        [ValidateRange(-90, 90)]
        [int]$TextRotation ,
        [Alias("AutoFit")]
        [Switch]$AutoSize,
        [float]$Width,
        [float]$Height,
        [Alias('Hide')]
        [Switch]$Hidden,
        [Switch]$Locked,
        [Switch]$Merge
    )
    process {
        if  ($Range -is [Array])  {
            $null = $PSBoundParameters.Remove("Range")
            $Range | Set-ExcelRange @PSBoundParameters
        }
        else {
            #We should accept, a worksheet and a name of a range or a cell address; a table; the address of a table; a named range; a row, a column or .Cells[ ]
            if ($Range -is [OfficeOpenXml.Table.ExcelTable]) {$Range = $Range.Address}
            elseif ($Worksheet -and ($Range -is [string] -or $Range -is [OfficeOpenXml.ExcelAddress])) {
                $Range = $Worksheet.Cells[$Range]
            }
            elseif ($Range -is [string]) {Write-Warning -Message "The range parameter you have specified also needs a worksheet parameter." ;return}
            #else we assume $Range is a range.
            if ($ClearAll)  {
                $Range.Clear()
            }
            if ($ResetFont) {
                $Range.Style.Font.Color.SetColor( ([System.Drawing.Color]::Black))
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.Bold          = $false
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.Italic        = $false
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.UnderLine     = $false
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.Strike        = $false
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
                $Range.Style.Font.VerticalAlign = [OfficeOpenXml.Style.ExcelVerticalAlignmentFont]::None
            }
            if ($PSBoundParameters.ContainsKey('Underline')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.UnderLine      = [boolean]$Underline
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.UnderLineType  = $UnderLineType
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('Bold')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.Bold           = [boolean]$bold
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('Italic')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.Italic         = [boolean]$italic
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('StrikeThru')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.Strike         = [boolean]$StrikeThru
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('FontSize')){
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.Size           = $FontSize
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('FontName')){
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.Name           = $FontName
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('FontShift')){
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Font.VerticalAlign  = $FontShift
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('FontColor')){
                if ($FontColor -is [string]) {$FontColor = [System.Drawing.Color]::$FontColor }
                $Range.Style.Font.Color.SetColor(  $FontColor)
            }
            if ($PSBoundParameters.ContainsKey('TextRotation')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.TextRotation        = $TextRotation
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('WrapText')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.WrapText            = [boolean]$WrapText
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('HorizontalAlignment')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.HorizontalAlignment = $HorizontalAlignment
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('VerticalAlignment')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.VerticalAlignment   = $VerticalAlignment
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($PSBoundParameters.ContainsKey('Merge')) {
                $Range.Merge   = [boolean]$Merge
            }
            if ($PSBoundParameters.ContainsKey('Value')) {
                if ($Value -match '^=')      {$PSBoundParameters["Formula"] = $Value -replace '^=','' }
                else {
                    $Range.Value = $Value
                    try {  # HA-LINUX-PATCH-v1
                        if ($Value -is [datetime])  { $Range.Style.Numberformat.Format = 'm/d/yy h:mm' }
                    } catch {
                        Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                    }
                    try {  # HA-LINUX-PATCH-v1
                        if ($Value -is [timespan])  { $Range.Style.Numberformat.Format = '[h]:mm:ss'   }
                    } catch {
                        Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                    }
                }
            }
            if ($PSBoundParameters.ContainsKey('Formula')) {
                if ($ArrayFormula) {$Range.CreateArrayFormula(($Formula -replace '^=','')) }
                else               {$Range.Formula         =  ($Formula -replace '^=','')  }
            }
            if ($PSBoundParameters.ContainsKey('NumberFormat')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Numberformat.Format = (Expand-NumberFormat $NumberFormat)
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
            if ($BorderColor -is [string]) {$BorderColor = [System.Drawing.Color]::$BorderColor }
            if ($PSBoundParameters.ContainsKey('BorderAround')) {
                $Range.Style.Border.BorderAround($BorderAround, $BorderColor)
            }
            if ($PSBoundParameters.ContainsKey('BorderBottom')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Border.Bottom.Style=$BorderBottom
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
                $Range.Style.Border.Bottom.Color.SetColor($BorderColor)
            }
            if ($PSBoundParameters.ContainsKey('BorderTop')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Border.Top.Style=$BorderTop
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
                $Range.Style.Border.Top.Color.SetColor($BorderColor)
            }
            if ($PSBoundParameters.ContainsKey('BorderLeft')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Border.Left.Style=$BorderLeft
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
                $Range.Style.Border.Left.Color.SetColor($BorderColor)
            }
            if ($PSBoundParameters.ContainsKey('BorderRight')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Border.Right.Style=$BorderRight
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
                $Range.Style.Border.Right.Color.SetColor($BorderColor)
            }
            if ($PSBoundParameters.ContainsKey('BackgroundColor')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Fill.PatternType = $BackgroundPattern
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
                if ($BackgroundColor -is [string]) {$BackgroundColor = [System.Drawing.Color]::$BackgroundColor }
                $Range.Style.Fill.BackgroundColor.SetColor($BackgroundColor)
                if ($PatternColor) {
                    if ($PatternColor -is [string]) {$PatternColor = [System.Drawing.Color]::$PatternColor }
                    $Range.Style.Fill.PatternColor.SetColor( $PatternColor)
                }
            }
            if ($PSBoundParameters.ContainsKey('Height')) {
                if ($Range -is [OfficeOpenXml.ExcelRow]   ) {$Range.Height = $Height }
                elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                    ($Range.Start.Row)..($Range.Start.Row + $Range.Rows) |
                        ForEach-Object {$Range.Worksheet.Row($_).Height = $Height }
                }
                else {Write-Warning -Message ("Can set the height of a row or a range but not a {0} object" -f ($Range.GetType().name)) }
            }
            if ($Autosize -and -not $env:NoAutoSize) {
                try {
                    if ($Range -is [OfficeOpenXml.ExcelColumn]) {$Range.AutoFit() }
                    elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                        $Range.AutoFitColumns()

                    }
                    else {Write-Warning -Message ("Can autofit a column or a range but not a {0} object" -f ($Range.GetType().name)) }
                }
                catch {Write-Warning -Message "Failed autosizing columns of worksheet '$WorksheetName': $_"}
            }
            elseif ($AutoSize) {Write-Warning -Message "Auto-fitting columns is not available with this OS configuration." }
            elseif ($PSBoundParameters.ContainsKey('Width')) {
                if ($Range -is [OfficeOpenXml.ExcelColumn]) {$Range.Width = $Width}
                elseif ($Range -is [OfficeOpenXml.ExcelRange] ) {
                    ($Range.Start.Column)..($Range.Start.Column + $Range.Columns - 1) |
                        ForEach-Object {
                            #$ws.Column($_).Width = $Width
                            $Range.Worksheet.Column($_).Width = $Width
                        }
                }
                else {Write-Warning -Message ("Can set the width of a column or a range but not a {0} object" -f ($Range.GetType().name)) }
            }
            if ($PSBoundParameters.ContainsKey('Hidden')) {
                if ($Range -is [OfficeOpenXml.ExcelRow] -or
                    $Range -is [OfficeOpenXml.ExcelColumn]  ) {$Range.Hidden = [boolean]$Hidden}
                else {Write-Warning -Message ("Can hide a row or a column but not a {0} object" -f ($Range.GetType().name)) }
            }
            if ($PSBoundParameters.ContainsKey('Locked')) {
                try {  # HA-LINUX-PATCH-v1: skip on EPPlus/Linux runtime where Style setters may be missing
                    $Range.Style.Locked=$Locked
                } catch {
                    Write-Verbose "[HA-LinuxPatch] Skipping Style assignment on Linux: $((($_.Exception.Message) -replace '\r?\n',' ' | Select-Object -First 1))"
                }
            }
        }
    }
}
