function Start-M365DSCConfigurationExtract
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param(
        [Parameter()]
        [Switch]
        $Quiet,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount,

        [Parameter()]
        [System.String[]]
        $ComponentsToExtract,

        [Parameter()]
        [Switch]
        $AllComponents,

        [Parameter()]
        [System.String]
        $Path,

        [Parameter()]
        [System.String]
        $FileName,

        [Parameter()]
        [System.String]
        $ConfigurationName = 'M365TenantConfig',

        [Parameter()]
        [ValidateRange(1, 100)]
        $MaxProcesses = 16,

        [Parameter()]
        [ValidateSet('AAD', 'SPO', 'EXO', 'SC', 'OD', 'O365', 'TEAMS', 'PP')]
        [System.String[]]
        $Workloads,

        [Parameter()]
        [ValidateSet('Lite', 'Default', 'Full')]
        [System.String]
        $Mode,

        [Parameter()]
        [System.Boolean]
        $GenerateInfo = $false,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.string]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword
    )
    $M365DSCExportStartTime = [System.DateTime]::Now
    $InformationPreference = "Continue"

    Reset-AllTeamsCached

    $DefaultWarningPreference = $WarningPreference
    $DefaultVerbosePreference = $VerbosePreference

    # We will set this on our own in app, it is already on SiletntlyContinue by default,
    # but we dont want it here explicitly overriding the variable
    # $VerbosePreference = "SilentlyContinue"
    # $WarningPreference = "SilentlyContinue"

    if ($null -eq $ComponentsToExtract -or $ComponentsToExtract.Length -eq 0)
    {
        $ComponentsToExtractSpecified = $false
    }
    else
    {
        $ComponentsToExtractSpecified = $true
    }
    $organization = ""
    $principal = "" # Principal represents the "NetBios" name of the tenant (e.g. the M365DSC part of M365DSC.onmicrosoft.com)
    $ConnectionMode = $null
    if (-not [String]::IsNullOrEmpty($ApplicationId) -and `
            -not [String]::IsNullOrEmpty($TenantId) -and `
            -not [String]::IsNullOrEmpty($CertificateThumbprint))
    {
        $ConnectionMode = 'ServicePrincipal'
        $organization = Get-M365DSCTenantDomain -ApplicationId $ApplicationId `
            -TenantId $TenantId `
            -CertificateThumbprint $CertificateThumbprint
    }
    elseif (-not [String]::IsNullOrEmpty($CertificatePath))
    {
        $ConnectionMode = 'ServicePrincipal'
        $organization = $TenantId
    }
    elseif (-not [String]::IsNullOrEmpty($GlobalAdminAccount))
    {
        $ConnectionMode = 'Credential'
        if ($null -ne $GlobalAdminAccount -and $GlobalAdminAccount.UserName.Contains("@"))
        {
            $organization = $GlobalAdminAccount.UserName.Split("@")[1]
        }
    }
    if ($organization.IndexOf(".") -gt 0)
    {
        $principal = $organization.Split(".")[0]
    }

    $ComponentsToSkip = @()
    if ($Mode -eq 'Default')
    {
        $ComponentsToSkip = $Global:FullComponents
    }
    elseif ($Mode -eq 'Lite')
    {
        $ComponentsToSkip = $Global:DefaultComponents + $Global:FullComponents
    }

    $AzureAutomation = $false
    $version = (Get-Module 'Microsoft365DSC').Version

    if ($ConnectionMode -eq 'Credential')
    {
        # Add the GlobalAdminAccount to the Credentials List
        Save-Credentials -UserName "globaladmin"
    }
    else
    {
        Save-Credentials -UserName "certificatepassword"
    }

    if ($ConnectionMode -ne 'Credential')
    {
        Add-ConfigurationDataEntry -Node "NonNodeData" `
            -Key "OrganizationName" `
            -Value $organization `
            -Description "Tenant's default verified domain name"
        Add-ConfigurationDataEntry -Node "NonNodeData" `
            -Key "ApplicationId" `
            -Value $ApplicationId `
            -Description "Azure AD Application Id for Authentication"
        if (-not [System.String]::IsNullOrEmpty($TenantId))
        {
            Add-ConfigurationDataEntry -Node "NonNodeData" `
                -Key "TenantId" `
                -Value $TenantId `
                -Description "The Id or Name of the tenant to authenticate against"
        }

        if (-not [System.String]::IsNullOrEmpty($CertificatePath))
        {
            Add-ConfigurationDataEntry -Node "NonNodeData" `
                -Key "CertificatePath" `
                -Value $CertificatePath `
                -Description "Local path to the .pfx certificate to use for authentication"
        }

        if (-not [System.String]::IsNullOrEmpty($CertificateThumbprint))
        {
            Add-ConfigurationDataEntry -Node "NonNodeData" `
                -Key "CertificateThumbprint" `
                -Value $CertificateThumbprint `
                -Description "Thumbprint of the certificate to use for authentication"
        }
    }

    Add-ConfigurationDataEntry -Node "localhost" `
        -Key "ServerNumber" `
        -Value "0" `
        -Description "Default Value Used to Ensure a Configuration Data File is Generated"


    $shouldOpenOutputDirectory = !$Quiet
    #region Prompt the user for a location to save the extract and generate the files
    if ([System.String]::IsNullOrEmpty($Path))
    {
        $shouldOpenOutputDirectory = $true
        $OutputDSCPath = Read-Host "Destination Path"
    }
    else
    {
        $OutputDSCPath = $Path
    }

    if ([System.String]::IsNullOrEmpty($OutputDSCPath))
    {
        $OutputDSCPath = '.'
    }

    while ((Test-Path -Path $OutputDSCPath -PathType Container -ErrorAction SilentlyContinue) -eq $false)
    {
        try
        {
            Write-Information "Directory `"$OutputDSCPath`" doesn't exist; creating..."
            New-Item -Path $OutputDSCPath -ItemType Directory | Out-Null
            if ($?)
            {
                break
            }
        }
        catch
        {
            Write-Warning "$($_.Exception.Message)"
            Write-Warning "Could not create folder $OutputDSCPath!"
        }
        $OutputDSCPath = Read-Host "Please Provide Output Folder for DSC Configuration (Will be Created as Necessary)"
    }
    <## Ensures the path we specify ends with a Slash, in order to make sure the resulting file path is properly structured. #>
    if (!$OutputDSCPath.EndsWith("\") -and !$OutputDSCPath.EndsWith("/"))
    {
        $OutputDSCPath += "\"
    }
    #endregion


    # this is to avoid problems with unicode charachters when executing the generated ps1 file
    $Utf8BomEncoding = New-Object System.Text.UTF8Encoding $True




    $ResourcesPath = Join-Path -Path $PSScriptRoot `
        -ChildPath "..\DSCResources\" `
        -Resolve
    $AllResources = Get-ChildItem $ResourcesPath -Recurse | Where-Object { $_.Name -like 'MSFT_*.psm1' }




    $i = 1
    $ResourcesToExport = @()
    $resourceExtractionStates = @{}
    foreach ($ResourceModule in $AllResources)
    {
        try
        {
            $msftResourceName = $ResourceModule.Name.Split('.')[0];
            $resourceName = $msftResourceName.Replace('MSFT_', '')
            [array]$currentWorkload = $ResourceName.Substring(0, 2)
            switch ($currentWorkload.ToUpper())
            {
                'AA'
                {
                    $currentWorkload = 'AAD';
                    break
                }
                'EX'
                {
                    $currentWorkload = 'EXO';
                    break
                }
                'O3'
                {
                    $currentWorkload = 'O365';
                    break
                }
                'OD'
                {
                    $currentWorkload = 'OD';
                    break
                }
                'PL'
                {
                    $currentWorkload = 'PLANNER';
                    break
                }
                'PP'
                {
                    $currentWorkload = 'PP';
                    break
                }
                'SC'
                {
                    $currentWorkload = 'SC';
                    break
                }
                'SP'
                {
                    $currentWorkload = 'SPO';
                    break
                }
                'TE'
                {
                    $currentWorkload = 'Teams';
                    break
                }
                default
                {
                    $currentWorkload = $null;
                    break
                }
            }

            $resourceExtractionStates[$msftResourceName] = 'NotIncluded'
            if (($null -ne $ComponentsToExtract -and
                    ($ComponentsToExtract -contains $resourceName -or $ComponentsToExtract -contains ("chck" + $resourceName))) -or
                $AllComponents -or ($null -ne $Workloads -and $Workloads -contains $currentWorkload) -or `
                ($null -eq $ComponentsToExtract -and $null -eq $Workloads) -and `
                ($ComponentsToExtractSpecified -or -not $ComponentsToSkip.Contains($resourceName)))
            {
                $ResourcesToExport += $ResourceModule
                $resourceExtractionStates[$msftResourceName] = 'ToBeLoaded'
            }
        }
        catch
        {
            New-M365DSCLogEntry -Error $_ -Message $ResourceModule.Name -Source "[M365DSCReverse]$($ResourceModule.Name)"
        }
    }

    $platformSkipsNotified = @()
    $resourceTimeTotalTaken = @{}
    foreach ($resource in $ResourcesToExport)
    {
        $stopWatch = [system.diagnostics.stopwatch]::StartNew()
        $msftResourceName = $resource.Name.Split('.')[0];
        $resourceName = $msftResourceName.Replace('MSFT_', '')
        try
        {
            $shouldSkipBecauseOfFailedPlatforms = $false
            [array]$usedPlatforms = Get-ResourcePlatformUsage -Resource $resourceName -ResourceModuleFilePath $resource.FullName
            foreach ($platform in $usedPlatforms)
            {
                # we will skip PnP if there was a problem connecting to a specific site
                # it could be a permissions issue
                # if it was a problem with connecting to the admin site, then we know that all else will fail as well so no need to continue
                if ($platform -eq 'PnP' -and $null -ne $Global:SPOAdminUrl -and $Global:SPOConnectionUrl -ne $Global:SPOAdminUrl)
                {
                    continue
                }
                $isAvailable = Check-PlatformAvailability -Platform $platform
                Write-Verbose "The isConnectionAvailable flag is $isAvailable for $platform"
                $shouldSkipBecauseOfFailedPlatforms = $shouldSkipBecauseOfFailedPlatforms -or !$isAvailable
                Write-Verbose "The shouldskip flag is $shouldSkipBecauseOfFailedPlatforms for $platform"
                if (!$isAvailable -and !$platformSkipsNotified.Contains($platform))
                {
                    Write-Error "The [$platform] connection has failed and all of the related resources will be skipped to avoid unnecessary errors."
                    $platformSkipsNotified += $platform
                }
            }

            if ($shouldSkipBecauseOfFailedPlatforms)
            {
                $resourceExtractionStates[$msftResourceName] = 'SkippedBecauseOfPreviousConnectionFailure'
                Write-Verbose "Skipped [$resourceName] because of connection problems with the used MsCloudLogin platform"
                continue;
            }

            Import-Module $resource.FullName | Out-Null
            $MaxProcessesExists = (Get-Command 'Export-TargetResource').Parameters.Keys.Contains("MaxProcesses")
            $AppSecretExists = (Get-Command 'Export-TargetResource').Parameters.Keys.Contains("ApplicationSecret")
            $CertThumbprintExists = (Get-Command 'Export-TargetResource').Parameters.Keys.Contains("CertificateThumbprint")
            $TenantIdExists = (Get-Command 'Export-TargetResource').Parameters.Keys.Contains("TenantId")
            $AppIdExists = (Get-Command 'Export-TargetResource').Parameters.Keys.Contains("ApplicationId")
            $GlobalAdminExists = (Get-Command 'Export-TargetResource').Parameters.Keys.Contains("GlobalAdminAccount")
            $CertPathExists = (Get-Command 'Export-TargetResource').Parameters.Keys.Contains("CertificatePath")
            $CertPasswordExists = (Get-Command 'Export-TargetResource').Parameters.Keys.Contains("CertificatePassword")

            $parameters = @{}
            if ($GlobalAdminExists -and -not [System.String]::IsNullOrEmpty($GlobalAdminAccount))
            {
                $parameters.Add("GlobalAdminAccount", $GlobalAdminAccount)
            }
            if ($MaxProcessesExists -and -not [System.String]::IsNullOrEmpty($MaxProcesses))
            {
                $parameters.Add("MaxProcesses", $MaxProcesses)
            }
            if ($AppSecretExists -and -not [System.String]::IsNullOrEmpty($ApplicationSecret))
            {
                $parameters.Add("AppplicationSecret", $ApplicationSecret)
            }
            if ($CertThumbprintExists -and -not [System.String]::IsNullOrEmpty($CertificateThumbprint))
            {
                $parameters.Add("CertificateThumbprint", $CertificateThumbprint)
            }
            if ($TenantIdExists -and -not [System.String]::IsNullOrEmpty($TenantId))
            {
                $parameters.Add("TenantId", $TenantId)
            }
            if ($AppIdExists -and -not [System.String]::IsNullOrEmpty($ApplicationId))
            {
                $parameters.Add("ApplicationId", $ApplicationId)
            }
            if ($CertPathExists -and -not [System.String]::IsNullOrEmpty($CertificatePath))
            {
                $parameters.Add("CertificatePath", $CertificatePath)
            }
            if ($CertPasswordExists -and $null -ne $CertificatePassword)
            {
                $parameters.Add("CertificatePassword", $CertificatePassword)
            }
            if ($ComponentsToSkip -notcontains $resourceName)
            {
                Write-Host "[$i/$($ResourcesToExport.Length)] Extracting [$($resource.Name.Split('.')[0].Replace('MSFT_', ''))]..." -NoNewline
                $exportString = ""
                if ($GenerateInfo)
                {
                    $exportString += "`r`n        # For information on how to use this resource, please refer to:`r`n"
                    $exportString += "        # https://github.com/microsoft/Microsoft365DSC/wiki/$resourceName`r`n"
                }
                $exportString += Export-TargetResource @parameters
                $i++

                $psParseErrorOccurred = $false
                try
                {
                    [System.Management.Automation.Language.Token[]]$tokens = $null;
                    [System.Management.Automation.Language.ParseError[]]$parseErrors = $null;
                    [void][System.Management.Automation.Language.Parser]::ParseInput($exportString, [ref]$tokens, [ref]$parseErrors)
                    if ($parseErrors.Length -gt 0)
                    {
                        Write-Error "The [$resourceName] resource encountered an error and will not be available in the extracted data"
                        Write-Verbose "The [$resourceName] had parse errors"
                        Write-Verbose "#######PARSE ERRORS START####################"
                        foreach ($parseError in $parseErrors)
                        {
                            Write-Verbose $parseError
                        }
                        Write-Verbose "#######PARSE INPUT WITH ERRORS START######################"
                        Write-Verbose $exportString
                        Write-Verbose "#######PARSE ERRORS END#############################"
                        $psParseErrorOccurred = $true
                    }
                }
                catch
                {
                    $psParseErrorOccurred = $true
                    Write-Error $_
                }

                if ($psParseErrorOccurred)
                {
                    $exportString = ""
                    $resourceExtractionStates[$msftResourceName] = 'PsParseError'
                }
                else
                {
                    $resourceExtractionStates[$msftResourceName] = 'Extracted'
                }
            }
            else
            {
                $resourceExtractionStates[$msftResourceName] = 'NotIncluded'
            }

            $fileStream = $null
            $sw = $null
            try
            {
                $resOutputFilePath = Join-Path $OutputDSCPath "$($resourceName)_TenantConfig.ps1"
                $fileStream = [System.IO.File]::OpenWrite("$resOutputFilePath")

                $sw = New-Object System.IO.StreamWriter -ArgumentList @($fileStream, $Utf8BomEncoding)
                Write-DscStartFileContents -Writer $sw
                $sw.Write($exportString)
                Write-DscEndingFileContents $sw
            }
            finally
            {
                if ($sw)
                {
                    $sw.Dispose()
                }
                if ($fileStream)
                {
                    $fileStream.Dispose()
                }
            }

            $exportString = $null
        }
        catch
        {
            $resourceExtractionStates[$msftResourceName] = 'ExtractionError'
            $ex = $_.Exception
            $isMissingGrantError = $false

            while ($null -ne $ex -and $ex.GetType().FullName -ne "Microsoft.IdentityModel.Clients.ActiveDirectory.AdalServiceException")
            {
                $ex = $ex.InnerException
            }

            if ($null -ne $ex -and $ex.ErrorCode -eq "invalid_grant" -and $null -ne $ex.ServiceErrorCodes -and $ex.ServiceErrorCodes.Contains("65001"))
            {
                $isMissingGrantError = $true
            }

            if ($isMissingGrantError -and $currentWorkload -eq 'PP')
            {
                Write-Verbose "PowerApps service app permissions are not granted. The enteriprise application is most likely missing.`nVisit the PowerApps admin center to get it created and rerun the Trace Configuration Wizard"

                #don't want any messages in the UI
                $platformSkipsNotified += "PowerPlatforms"
            }
            else
            {
                New-M365DSCLogEntry -Error $_ -Message $ResourceModule.Name -Source "[O365DSCReverse]$($ResourceModule.Name)"
            }
        }
        finally
        {
            $stopWatch.Stop()
            $WarningPreference = $DefaultWarningPreference;
            $VerbosePreference = $DefaultVerbosePreference;
        }

        $resourceTimeTotalTaken[$msftResourceName] = $stopWatch.Elapsed.TotalSeconds
    }


    # dont' leave dangling remote sessions, there can only be a couple of them for EXO and SC
    Remove-RemoteSessions

    #region Benchmarks
    $M365DSCExportEndTime = [System.DateTime]::Now
    $timeTaken = New-Timespan -Start ($M365DSCExportStartTime.ToString()) `
        -End ($M365DSCExportEndTime.ToString())
    Write-Host "$($Global:M365DSCEmojiHourglass) Export took {" -NoNewLine
    Write-Host "$($timeTaken.TotalSeconds) seconds" -NoNewLine -ForegroundColor Cyan
    Write-Host "}"
    #endregion

    if (-not [System.String]::IsNullOrEmpty($FileName))
    {
        $outputDSCFile = $OutputDSCPath + $FileName
    }
    else
    {
        $outputDSCFile = $OutputDSCPath + "M365TenantConfig.ps1"
    }

    Write-ExtractionStates -OutputDSCPath $OutputDSCPath -ResourceExtractionStates $resourceExtractionStates -ResourceTimeTotalTaken $resourceTimeTotalTaken

    if (!$AzureAutomation)
    {
        $LCMConfig = Get-DscLocalConfigurationManager
        if ($null -ne $LCMConfig.CertificateID)
        {
            try
            {
                # Export the certificate assigned to the LCM
                $certPath = $OutputDSCPath + "M365DSC.cer"
                Export-Certificate -FilePath $certPath `
                    -Cert "cert:\LocalMachine\my\$($LCMConfig.CertificateID)" `
                    -Type CERT `
                    -NoClobber | Out-Null
                Add-ConfigurationDataEntry -Node "localhost" `
                    -Key "CertificateFile" `
                    -Value "M365DSC.cer" `
                    -Description "Path of the certificate used to encrypt credentials in the file."
            }
            catch
            {
                Write-Verbose -Message $_
            }
        }
        $outputConfigurationData = $OutputDSCPath + "ConfigurationData.psd1"
        New-ConfigurationDataDocument -Path $outputConfigurationData
    }

    if ($shouldOpenOutputDirectory)
    {
        Invoke-Item -Path $OutputDSCPath
    }
}


function Remove-RemoteSessions
{
    $prevConfirmPreference = $ConfirmPreference
    try
    {
        $ConfirmPreference = 'None'

        # this will disconnect all of the Exhange but also all of the Security and Compliance sessions
        #Disconnect-ExchangeOnline

        # this code is basically copy pasted from Disconnect-ExchangeOnline
        # the reason why we are not using Disconnect-ExchangeOnline is because it asks for a confirmation and $ConfirmPreference = 'None' does nothing, at least when running fron inside the script
        # inside a powershell window $ConfirmPreference = 'None' works as expected, but not when running within Trace or debugging in VSCode
        $existingPSSession = Get-PSSession | Where-Object { ($_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*") -or $_.Name -like 'SfBPowerShellSession*' }

        if ($existingPSSession.count -gt 0)
        {
            for ($index = 0; $index -lt $existingPSSession.count; $index++)
            {
                $session = $existingPSSession[$index]
                Remove-PSSession -session $session

                # Remove any previous modules loaded because of the current PSSession
                if ($null -ne $session.PreviousModuleName)
                {
                    if ((Get-Module $session.PreviousModuleName).Count -ne 0)
                    {
                        Remove-Module -Name $session.PreviousModuleName -ErrorAction SilentlyContinue
                    }

                    $session.PreviousModuleName = $null
                }

                # Remove any leaked module in case of removal of broken session object
                if ($null -ne $session.CurrentModuleName)
                {
                    if ((Get-Module $session.CurrentModuleName).Count -ne 0)
                    {
                        Remove-Module -Name $session.CurrentModuleName -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }
    catch
    {
        Write-Error "Error while disconnecting remote sessions"
        Write-Error $_
    }
    finally
    {
        $ConfirmPreference = $prevConfirmPreference
    }
}

function Write-ExtractionStates
{
    param(

        [Parameter(Mandatory = $true)]
        $OutputDSCPath,

        [Parameter(Mandatory = $true)]
        $ResourceExtractionStates,

        [Parameter(Mandatory = $true)]
        $ResourceTimeTotalTaken
    )

    # create a file containing the extraction states for all resources
    $outputExtractionStatesFile = Join-Path $OutputDSCPath "ExtractionStates.json"

    $resultingExtractionStatesObject = @{}
    foreach ($key in $ResourceExtractionStates.Keys)
    {
        $timeTaken = 0
        if ($ResourceTimeTotalTaken.ContainsKey($key))
        {
            $timeTaken = $ResourceTimeTotalTaken[$key]
        }
        $resultingExtractionStatesObject[$key] = @{
            State     = $ResourceExtractionStates[$key]
            TimeTaken = $timeTaken
        }
    }

    ConvertTo-Json -InputObject $resultingExtractionStatesObject | Out-File -FilePath $outputExtractionStatesFile -Encoding utf8
}

function Check-PlatformAvailability
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $Platform
    )

    $faulted = Get-Variable -Scope Global "MSCloudLogin${Platform}ConnectionFaulted" -ValueOnly -ErrorAction SilentlyContinue
    return $null -eq $faulted -or $faulted -eq $false
}

function Get-ResourcePlatformUsage
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $Resource,

        [Parameter(Mandatory = $true)]
        [string]
        $ResourceModuleFilePath
    )
    $fileContent = Get-Content $ResourceModuleFilePath -Raw
    $matches = [Regex]::Matches($fileContent, '-Platform\s+''?(?<platform>\w+)''?', [ System.Text.RegularExpressions.RegexOptions]::IgnoreCase);

    $platforms = @()
    foreach ($match in $matches)
    {
        $platform = $match.Groups["platform"].Value
        if ($platforms.Contains($platform))
        {
            continue
        }
        $platforms += $platform
    }

    return $platforms
}

function Write-DscStartFileContents
{
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.TextWriter]
        $Writer
    )
    $Writer.WriteLine("# Generated with Microsoft365DSC version $version")
    $Writer.WriteLine("# For additional information on how to use Microsoft365DSC, please visit https://aka.ms/M365DSC")
    if ($ConnectionMode -eq 'Credential')
    {
        $Writer.WriteLine("param (")
        $Writer.WriteLine("    [parameter()]")
        $Writer.WriteLine("    [System.Management.Automation.PSCredential]")
        $Writer.WriteLine("    `$GlobalAdminAccount")
        $Writer.WriteLine(")`r`n")
    }
    else
    {
        if (-not [System.String]::IsNullOrEmpty($CertificatePassword))
        {
            $Writer.WriteLine("param (")
            $Writer.WriteLine("    [parameter()]")
            $Writer.WriteLine("    [System.Management.Automation.PSCredential]")
            $Writer.WriteLine("    `$CertificatePassword")
            $Writer.WriteLine(")`r`n")
        }
    }

    $ConfigurationName = 'M365TenantConfig'

    $Writer.WriteLine("Configuration $ConfigurationName`r`n{")

    if ($ConnectionMode -eq 'Credential')
    {
        $Writer.WriteLine("    param (")
        $Writer.WriteLine("        [parameter()]")
        $Writer.WriteLine("        [System.Management.Automation.PSCredential]")
        $Writer.WriteLine("        `$GlobalAdminAccount")
        $Writer.WriteLine("    )`r`n")
        $Writer.WriteLine("    if (`$null -eq `$GlobalAdminAccount)")
        $Writer.WriteLine("    {")
        $Writer.WriteLine("        <# Credentials #>")
        $Writer.WriteLine("    }")
        $Writer.WriteLine("    else")
        $Writer.WriteLine("    {")
        $Writer.WriteLine("        `$Credsglobaladmin = `$GlobalAdminAccount")
        $Writer.WriteLine("    }`r`n")
        $Writer.WriteLine("    `$OrganizationName = `$Credsglobaladmin.UserName.Split('@')[1]")
    }
    else
    {
        if (-not [System.String]::IsNullOrEmpty($CertificatePassword))
        {
            $Writer.WriteLine("    param (")
            $Writer.WriteLine("        [parameter()]")
            $Writer.WriteLine("        [System.Management.Automation.PSCredential]")
            $Writer.WriteLine("        `$CertificatePassword")
            $Writer.WriteLine("    )`r`n")
            $Writer.WriteLine("    if (`$null -eq `$CertificatePassword)")
            $Writer.WriteLine("    {")
            $Writer.WriteLine("        <# Credentials #>")
            $Writer.WriteLine("    }")
            $Writer.WriteLine("    else")
            $Writer.WriteLine("    {")
            $Writer.WriteLine("        `$CredsCertificatePassword = `$CertificatePassword")
            $Writer.WriteLine("    }`r`n")
        }

        $Writer.WriteLine("    `$OrganizationName = `$ConfigurationData.NonNodeData.OrganizationName")
    }
    $Writer.WriteLine("    Import-DscResource -ModuleName Microsoft365DSC`r`n")
    $Writer.WriteLine("    Node localhost")
    $Writer.WriteLine("    {")
}

function Write-DscEndingFileContents
{
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.TextWriter]
        $Writer
    )

    $ConfigurationName = 'M365TenantConfig'

    # Close the Node and Configuration declarations
    $writer.WriteLine("    }")
    $writer.WriteLine("}")

    $writer.WriteLine("$ConfigurationName -ConfigurationData .\ConfigurationData.psd1")
}
