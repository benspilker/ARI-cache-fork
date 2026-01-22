<#
.Synopsis
Main module for Excel Report Building

.DESCRIPTION
This module is the main module for building the Excel Report.

.Link
https://github.com/microsoft/ARI/Modules/Private/0.MainFunctions/Start-ARIReporOrchestration.ps1

.COMPONENT
This PowerShell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola

#>
Function Start-ARIReporOrchestration {
    Param($ReportCache,
    $SecurityCenter,
    $File,
    $Quotas,
    $SkipPolicy,
    $SkipAdvisory,
    $Automation,
    $TableStyle,
    $IncludeCosts)

    Write-Progress -activity 'Azure Inventory' -Status "65% Complete." -PercentComplete 65 -CurrentOperation "Starting the Report Phase.."

    <############################################################## REPORT CREATION ###################################################################>

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Resource Reporting Cache.')
    try {
        Start-ARIExcelJob -ReportCache $ReportCache -TableStyle $TableStyle -File $File
    } catch {
        Write-Error "Error in Start-ARIExcelJob: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        throw
    }

    <############################################################## REPORT EXTRA DETAILS ###################################################################>

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Reporting Extra Details.')
    try {
        Start-ARIExcelExtraData -File $File
    } catch {
        Write-Error "Error in Start-ARIExcelExtraData: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        throw
    }

    <############################################################## EXTRA REPORTS ###################################################################>

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Default Data Reporting.')

    try {
        Start-ARIExtraReports -File $File -Quotas $Quotas -SecurityCenter $SecurityCenter -SkipPolicy $SkipPolicy -SkipAdvisory $SkipAdvisory -IncludeCosts $IncludeCosts -TableStyle $TableStyle
    } catch {
        Write-Error "Error in Start-ARIExtraReports: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        throw
    }

}