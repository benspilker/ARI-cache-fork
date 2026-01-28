<#
.Synopsis
Inventory for Azure Outages

.DESCRIPTION
Excel Sheet Name: Outages

.Link
https://github.com/microsoft/ARI/Modules/APIs/Outages.ps1

.COMPONENT
This powershell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 4.0.1
First Release Date: 25th Aug, 2024
Authors: Claudio Merola 

#>

<######## Default Parameters. Don't modify this ########>

param($SCPath, $Sub, $Intag, $Resources, $Retirements, $Task, $File, $SmaResources, $TableStyle, $Unsupported)

If ($Task -eq 'Processing') {

    <######### Insert the resource extraction here ########>

    $Outages = $Resources | Where-Object { $_.TYPE -eq 'Microsoft.ResourceHealth/events' -and $_.properties.description -like '*How can customers make incidents like this less impactful?*' }

    <######### Insert the resource Process here ########>

    if($Outages)
        {
            function Convert-PlainTextFromHtml {
                param([string]$Value)
                if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
                $decoded = [System.Net.WebUtility]::HtmlDecode($Value)
                $stripped = $decoded -replace '<[^>]+>', ' '
                $collapsed = ($stripped -replace '\s+', ' ').Trim()
                return $collapsed
            }
            $tmp = foreach ($1 in $Outages) {
                # Safely extract impacted subscriptions
                $ImpactedSubs = @()
                if ($null -ne $1.properties -and $null -ne $1.properties.impact -and $null -ne $1.properties.impact.impactedRegions) {
                    if ($null -ne $1.properties.impact.impactedRegions.impactedSubscriptions) {
                        $ImpactedSubs = $1.properties.impact.impactedRegions.impactedSubscriptions | Select-Object -Unique
                    }
                }
                # If no impacted subscriptions found, try to extract from outage ID or use empty array
                if ($ImpactedSubs.Count -eq 0) {
                    # Try to extract subscription ID from outage ID if possible
                    if ($null -ne $1.id -and $1.id -match '/subscriptions/([^/]+)') {
                        $ImpactedSubs = @($matches[1])
                    }
                }
                
                $ResUCount = 1

                $Data = $1.properties

                foreach ($Sub0 in $ImpactedSubs)
                    {
                        $sub1 = $SUB | Where-Object { $_.id -eq $Sub0 }

                        $StartTime = $Data.impactStartTime
                        $StartTime = [datetime]$StartTime
                        $StartTime = $StartTime.ToString("yyyy-MM-dd HH:mm")

                        $Mitigation = $Data.impactMitigationTime
                        $Mitigation = [datetime]$Mitigation
                        $Mitigation = $Mitigation.ToString("yyyy-MM-dd HH:mm")

                        # Safely handle impactedService - check if it's an array before accessing .count
                        $impactedServiceValue = $1.properties.impact.impactedService
                        if ($null -ne $impactedServiceValue) {
                            if ($impactedServiceValue -is [System.Array] -and $impactedServiceValue.Count -gt 1) {
                                $ImpactedService = $impactedServiceValue | ForEach-Object { $_ + ' ,' }
                            } else {
                                $ImpactedService = $impactedServiceValue
                            }
                        } else {
                            $ImpactedService = ''
                        }
                        $ImpactedService = [string]$ImpactedService
                        $ImpactedService = if ($ImpactedService -like '* ,*') { $ImpactedService -replace ".$" }else { $ImpactedService }

                        # Safely parse HTML description
                        $OutageDescription = ''
                        $SplitDescription = @('', '', '', '', '', '', '')
                        try {
                            $HTML = New-Object -Com 'HTMLFile'
                            $HTML.write([ref]$1.properties.description)
                            $OutageDescription = $Html.body.innerText
                            # Split description into sections
                            $SplitDescription = $OutageDescription.split('How can we make our incident communications more useful?').split('How can customers make incidents like this less impactful?').split('How are we making incidents like this less likely or less impactful?').split('How did we respond?').split('What went wrong and why?').split('What happened?')
                        } catch {
                            # If HTML parsing fails, use raw description
                            $OutageDescription = $1.properties.description
                            $SplitDescription = @('', $OutageDescription, '', '', '', '', '')
                        }
                        
                        # Safely extract split description sections with bounds checking
                        $whatHappened = ''
                        $whatWentWrong = ''
                        $howDidWeRespond = ''
                        $howMakingLessLikely = ''
                        $howCustomersCanMakeLessImpactful = ''
                        
                        if ($SplitDescription.Count -gt 1 -and $null -ne $SplitDescription[1]) {
                            $whatHappenedLines = $SplitDescription[1].Split([Environment]::NewLine)
                            if ($whatHappenedLines.Count -gt 1) { $whatHappened = $whatHappenedLines[1] }
                        }
                        if ($SplitDescription.Count -gt 2 -and $null -ne $SplitDescription[2]) {
                            $whatWentWrongLines = $SplitDescription[2].Split([Environment]::NewLine)
                            if ($whatWentWrongLines.Count -gt 1) { $whatWentWrong = $whatWentWrongLines[1] }
                        }
                        if ($SplitDescription.Count -gt 3 -and $null -ne $SplitDescription[3]) {
                            $howDidWeRespondLines = $SplitDescription[3].Split([Environment]::NewLine)
                            if ($howDidWeRespondLines.Count -gt 1) { $howDidWeRespond = $howDidWeRespondLines[1] }
                        }
                        if ($SplitDescription.Count -gt 4 -and $null -ne $SplitDescription[4]) {
                            $howMakingLessLikelyLines = $SplitDescription[4].Split([Environment]::NewLine)
                            if ($howMakingLessLikelyLines.Count -gt 1) { $howMakingLessLikely = $howMakingLessLikelyLines[1] }
                        }
                        if ($SplitDescription.Count -gt 5 -and $null -ne $SplitDescription[5]) {
                            $howCustomersCanMakeLessImpactfulLines = $SplitDescription[5].Split([Environment]::NewLine)
                            if ($howCustomersCanMakeLessImpactfulLines.Count -gt 1) { $howCustomersCanMakeLessImpactful = $howCustomersCanMakeLessImpactfulLines[1] }
                        }

                        $whatHappened = Convert-PlainTextFromHtml $whatHappened
                        $whatWentWrong = Convert-PlainTextFromHtml $whatWentWrong
                        $howDidWeRespond = Convert-PlainTextFromHtml $howDidWeRespond
                        $howMakingLessLikely = Convert-PlainTextFromHtml $howMakingLessLikely
                        $howCustomersCanMakeLessImpactful = Convert-PlainTextFromHtml $howCustomersCanMakeLessImpactful

                        $obj = @{
                            'ID'                                                                  = $1.id;
                            'Subscription'                                                        = $sub1.name;
                            'Outage ID'                                                           = $1.name;
                            'Event Type'                                                          = $Data.eventType;
                            'Status'                                                              = $Data.status;
                            'Event Level'                                                         = $Data.eventlevel;
                            'Title'                                                               = $Data.title;
                            'Impact Start Time'                                                   = $StartTime;
                            'Impact Mitigation Time'                                              = $Mitigation;
                            'Impacted Services'                                                   = $ImpactedService;
                            'What happened'                                                       = $whatHappened;
                            'What went wrong and why'                                             = $whatWentWrong;
                            'How did we respond'                                                  = $howDidWeRespond;
                            'How are we making incidents like this less likely or less impactful' = $howMakingLessLikely;
                            'How can customers make incidents like this less impactful'           = $howCustomersCanMakeLessImpactful;
                            'Resource U'                                                          = $ResUCount
                        }
                        $obj
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }
                }
            $tmp
        }
}

<######## Resource Excel Reporting Begins Here ########>

Else {
    <######## $SmaResources.(RESOURCE FILE NAME) ##########>

    if ($SmaResources) {
        # Safely get Resource U sum - handle cases where property might not exist
        $ResourceUSum = 0
        try {
            $ResourceUValues = $SmaResources | Select-Object -ExpandProperty 'Resource U' -ErrorAction SilentlyContinue
            if ($ResourceUValues) {
                $ResourceUSum = ($ResourceUValues | Measure-Object -Sum).Sum
            }
        } catch {
            # If property doesn't exist, use count as fallback
            $ResourceUSum = if ($SmaResources -is [System.Array]) { $SmaResources.Count } else { 1 }
        }
        $TableName = ('OutageTab_'+$ResourceUSum)

        $Style = @(
        New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0' -Range 'A:E'
        New-ExcelStyle -HorizontalAlignment Left -NumberFormat '0' -WrapText -Width 55 -Range 'F:F'
        New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0' -Range 'G:I'
        New-ExcelStyle -HorizontalAlignment Left -NumberFormat '0' -WrapText -Width 80 -Range 'J:N'
        )

        $Exc = New-Object System.Collections.Generic.List[System.Object]
        $Exc.Add('Subscription')
        $Exc.Add('Outage ID')
        $Exc.Add('Event Type')       
        $Exc.Add('Status')
        $Exc.Add('Event Level')
        $Exc.Add('Title')
        $Exc.Add('Impact Start Time')
        $Exc.Add('Impact Mitigation Time')
        $Exc.Add('Impacted Services')
        $Exc.Add('What happened')
        $Exc.Add('What went wrong and why')
        $Exc.Add('How did we respond')
        $Exc.Add('How are we making incidents like this less likely or less impactful')
        $Exc.Add('How can customers make incidents like this less impactful')

        [PSCustomObject]$SmaResources | 
        ForEach-Object { $_ } | Select-Object $Exc | 
        Export-Excel -Path $File -WorksheetName 'Outages' -AutoSize -TableName $TableName -MaxAutoSizeRows 100 -TableStyle $tableStyle -Numberformat '0' -Style $Style

    }
}
