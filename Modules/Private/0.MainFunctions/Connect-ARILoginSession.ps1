<#
.Synopsis
Azure Login Session Module for Azure Resource Inventory

.DESCRIPTION
This module is used to invoke the authentication process that is handle by Azure PowerShell.

.Link
https://github.com/microsoft/ARI/Modules/Private/0.MainFunctions/Connect-LoginSession.ps1

.COMPONENT
This powershell Module is part of Azure Resource Inventory (ARI)

.NOTES
Version: 3.6.0
First Release Date: 15th Oct, 2024
Authors: Claudio Merola

#>
function Connect-ARILoginSession {
    Param($AzureEnvironment, $TenantID, $SubscriptionID, $DeviceLogin, $AppId, $Secret, $CertificatePath, $Debug)
    $DebugPreference = 'silentlycontinue'
    $ErrorActionPreference = 'Continue'

    function Invoke-QuietAzIdentityLogging {
        param([scriptblock]$ScriptBlock)

        $prevEnabled = $env:AZURE_IDENTITY_LOGGING_ENABLED
        $prevLevel = $env:AZURE_IDENTITY_LOGGING_LEVEL
        $prevDebug = $DebugPreference
        $prevVerbose = $VerbosePreference
        $prevInfo = $InformationPreference
        $prevAzDebug = $null
        $prevAzAccountsDebug = $null
        $hadAzDebug = $PSDefaultParameterValues.ContainsKey('Az.*:Debug')
        $hadAzAccountsDebug = $PSDefaultParameterValues.ContainsKey('Az.Accounts.*:Debug')
        if ($hadAzDebug) { $prevAzDebug = $PSDefaultParameterValues['Az.*:Debug'] }
        if ($hadAzAccountsDebug) { $prevAzAccountsDebug = $PSDefaultParameterValues['Az.Accounts.*:Debug'] }
        try {
            $env:AZURE_IDENTITY_LOGGING_ENABLED = 'false'
            $env:AZURE_IDENTITY_LOGGING_LEVEL = 'warning'
            $DebugPreference = 'SilentlyContinue'
            $VerbosePreference = 'SilentlyContinue'
            $InformationPreference = 'SilentlyContinue'
            $PSDefaultParameterValues['Az.*:Debug'] = $false
            $PSDefaultParameterValues['Az.Accounts.*:Debug'] = $false
            & $ScriptBlock
        } finally {
            $DebugPreference = $prevDebug
            $VerbosePreference = $prevVerbose
            $InformationPreference = $prevInfo
            if ($hadAzDebug) { $PSDefaultParameterValues['Az.*:Debug'] = $prevAzDebug } else { $PSDefaultParameterValues.Remove('Az.*:Debug') | Out-Null }
            if ($hadAzAccountsDebug) { $PSDefaultParameterValues['Az.Accounts.*:Debug'] = $prevAzAccountsDebug } else { $PSDefaultParameterValues.Remove('Az.Accounts.*:Debug') | Out-Null }
            if ($null -ne $prevEnabled) { $env:AZURE_IDENTITY_LOGGING_ENABLED = $prevEnabled } else { Remove-Item Env:AZURE_IDENTITY_LOGGING_ENABLED -ErrorAction SilentlyContinue }
            if ($null -ne $prevLevel) { $env:AZURE_IDENTITY_LOGGING_LEVEL = $prevLevel } else { Remove-Item Env:AZURE_IDENTITY_LOGGING_LEVEL -ErrorAction SilentlyContinue }
        }
    }

    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Starting Connect-LoginSession function')
    Write-Host $AzureEnvironment -BackgroundColor Green
    $Context = Invoke-QuietAzIdentityLogging { Get-AzContext -ErrorAction SilentlyContinue }
    if (!$TenantID) {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Tenant ID not specified')
        write-host "Tenant ID not specified. Use -TenantID parameter if you want to specify directly. "
        write-host "Authenticating Azure"
        write-host ""

        if($DeviceLogin.IsPresent)
            {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Logging with Device Login')
                Invoke-QuietAzIdentityLogging { Connect-AzAccount -UseDeviceAuthentication -Environment $AzureEnvironment | Out-Null }
            }
        else
            {
                try 
                    {
                        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Editing Login Experience')
                        $AZConfigNewLogin = Get-AzConfig -LoginExperienceV2 -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                        if ($AZConfigNewLogin.value -eq 'On' )
                            {
                                Invoke-QuietAzIdentityLogging {
                                    Update-AzConfig -LoginExperienceV2 Off | Out-Null
                                    Connect-AzAccount -Environment $AzureEnvironment | Out-Null
                                    Update-AzConfig -LoginExperienceV2 On | Out-Null
                                }
                            }
                        else
                            {
                                Invoke-QuietAzIdentityLogging { Connect-AzAccount -Environment $AzureEnvironment | Out-Null }
                            }
                    }
                catch
                    {
                        Invoke-QuietAzIdentityLogging { Connect-AzAccount -Environment $AzureEnvironment | Out-Null }
                    }
            }
        write-host ""
        write-host ""
        $Tenants = Invoke-QuietAzIdentityLogging { Get-AzTenant -WarningAction SilentlyContinue -InformationAction SilentlyContinue | Sort-Object -Unique }
        if ($Tenants.Count -eq 1) {
            write-host "You have privileges only in One Tenant "
            write-host ""
            $TenantID = $Tenants.Id
        }
        else {
            write-host "Select the the Azure Tenant ID that you want to connect : "
            write-host ""
            $SequenceID = 1
            foreach ($Tenant in $Tenants) {
                $TenantName = $Tenant.name
                write-host "$SequenceID)  $TenantName"
                $SequenceID ++
            }
            write-host ""
            [int]$SelectTenant = read-host "Select Tenant ( default 1 )"
            $defaultTenant = --$SelectTenant
            $TenantID = ($Tenants[$defaultTenant]).Id
            if($DeviceLogin.IsPresent)
                {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Logging with Device Login')
                    Invoke-QuietAzIdentityLogging { Connect-AzAccount -Tenant $TenantID -UseDeviceAuthentication -Environment $AzureEnvironment | Out-Null }
                }
            else
                {
                    Invoke-QuietAzIdentityLogging { Connect-AzAccount -Tenant $TenantID -Environment $AzureEnvironment | Out-Null }
                }
        }
    }
    else {
        Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Tenant ID was informed.')

        if($Context.Tenant.Id -ne $TenantID)
        {
            Invoke-QuietAzIdentityLogging { Set-AzContext -Tenant $TenantID -ErrorAction SilentlyContinue | Out-Null }
            $Context = Invoke-QuietAzIdentityLogging { Get-AzContext -ErrorAction SilentlyContinue }
        }
        $Subs = Invoke-QuietAzIdentityLogging { Get-AzSubscription -TenantId $TenantID -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }

        if($DeviceLogin.IsPresent)
            {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Logging with Device Login')
                Invoke-QuietAzIdentityLogging { Connect-AzAccount -Tenant $TenantID -UseDeviceAuthentication -Environment $AzureEnvironment | Out-Null }
            }
        elseif($AppId -and $Secret -and $CertificatePath -and $TenantID)
            {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Logging with AppID and CertificatePath')
                $SecurePassword = ConvertTo-SecureString -String $Secret -AsPlainText -Force
                Invoke-QuietAzIdentityLogging { Connect-AzAccount -ServicePrincipal -TenantId $TenantId -ApplicationId $AppId -CertificatePath $CertificatePath -CertificatePassword $SecurePassword | Out-Null }
            }            
        elseif($AppId -and $Secret -and $TenantID)
            {
                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Logging with AppID and Secret')
                $SecurePassword = ConvertTo-SecureString -String $Secret -AsPlainText -Force
                $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppId, $SecurePassword
                Invoke-QuietAzIdentityLogging { Connect-AzAccount -ServicePrincipal -TenantId $TenantId -Credential $Credential | Out-Null }
            }
        else
            {
                if([string]::IsNullOrEmpty($Subs))
                    {
                        try 
                            {
                                Write-Debug ((get-date -Format 'yyyy-MM-dd_HH_mm_ss')+' - '+'Editing Login Experience')
                                $AZConfig = Get-AzConfig -LoginExperienceV2 -WarningAction SilentlyContinue -InformationAction SilentlyContinue
                                if ($AZConfig.value -eq 'On')
                                    {
                                        Invoke-QuietAzIdentityLogging {
                                            Update-AzConfig -LoginExperienceV2 Off | Out-Null
                                            Connect-AzAccount -Tenant $TenantID -Environment $AzureEnvironment | Out-Null
                                            Update-AzConfig -LoginExperienceV2 On | Out-Null
                                        }
                                    }
                                else
                                    {
                                        Invoke-QuietAzIdentityLogging { Connect-AzAccount -Tenant $TenantID -Environment $AzureEnvironment | Out-Null }
                                    }
                            }
                        catch
                            {
                                Invoke-QuietAzIdentityLogging { Connect-AzAccount -Tenant $TenantID -Environment $AzureEnvironment | Out-Null }
                            }
                    }
                else
                    {
                        Write-Host "Already authenticated in Tenant $TenantID"
                    }
            }
    }
    return $TenantID
}
