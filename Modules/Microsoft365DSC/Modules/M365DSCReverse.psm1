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
    $DSCContent = "# Generated with Microsoft365DSC version $version`r`n"
    $DSCContent += "# For additional information on how to use Microsoft365DSC, please visit https://aka.ms/M365DSC`r`n"
    if ($ConnectionMode -eq 'Credential')
    {
        $DSCContent += "param (`r`n"
        $DSCContent += "    [parameter()]`r`n"
        $DSCContent += "    [System.Management.Automation.PSCredential]`r`n"
        $DSCContent += "    `$GlobalAdminAccount`r`n"
        $DSCContent += ")`r`n`r`n"
    }
    else
    {
        if (-not [System.String]::IsNullOrEmpty($CertificatePassword))
        {
            $DSCContent += "param (`r`n"
            $DSCContent += "    [parameter()]`r`n"
            $DSCContent += "    [System.Management.Automation.PSCredential]`r`n"
            $DSCContent += "    `$CertificatePassword`r`n"
            $DSCContent += ")`r`n`r`n"
        }
    }

    if (-not [System.String]::IsNullOrEmpty($FileName))
    {
        $FileParts = $FileName.Split('.')

        if ([System.String]::IsNullOrEmpty($ConfigurationName))
        {
            $ConfigurationName = $FileName.Replace('.' + $FileParts[$FileParts.Length - 1], "")
        }
    }
    if ([System.String]::IsNullOrEmpty($ConfigurationName))
    {
        $ConfigurationName = 'M365TenantConfig'
    }
    $DSCContent += "Configuration $ConfigurationName`r`n{`r`n"

    if ($ConnectionMode -eq 'Credential')
    {
        $DSCContent += "    param (`r`n"
        $DSCContent += "        [parameter()]`r`n"
        $DSCContent += "        [System.Management.Automation.PSCredential]`r`n"
        $DSCContent += "        `$GlobalAdminAccount`r`n"
        $DSCContent += "    )`r`n`r`n"
        $DSCContent += "    if (`$null -eq `$GlobalAdminAccount)`r`n"
        $DSCContent += "    {`r`n"
        $DSCContent += "        <# Credentials #>`r`n"
        $DSCContent += "    }`r`n"
        $DSCContent += "    else`r`n"
        $DSCContent += "    {`r`n"
        $DSCContent += "        `$Credsglobaladmin = `$GlobalAdminAccount`r`n"
        $DSCContent += "    }`r`n`r`n"
        $DSCContent += "    `$OrganizationName = `$Credsglobaladmin.UserName.Split('@')[1]`r`n"
    }
    else
    {
        if (-not [System.String]::IsNullOrEmpty($CertificatePassword))
        {
            $DSCContent += "    param (`r`n"
            $DSCContent += "        [parameter()]`r`n"
            $DSCContent += "        [System.Management.Automation.PSCredential]`r`n"
            $DSCContent += "        `$CertificatePassword`r`n"
            $DSCContent += "    )`r`n`r`n"
            $DSCContent += "    if (`$null -eq `$CertificatePassword)`r`n"
            $DSCContent += "    {`r`n"
            $DSCContent += "        <# Credentials #>`r`n"
            $DSCContent += "    }`r`n"
            $DSCContent += "    else`r`n"
            $DSCContent += "    {`r`n"
            $DSCContent += "        `$CredsCertificatePassword = `$CertificatePassword`r`n"
            $DSCContent += "    }`r`n`r`n"
        }

        $DSCContent += "    `$OrganizationName = `$ConfigurationData.NonNodeData.OrganizationName`r`n"
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
    $DSCContent += "    Import-DscResource -ModuleName Microsoft365DSC`r`n`r`n"
    $DSCContent += "    Node localhost`r`n"
    $DSCContent += "    {`r`n"

    Add-ConfigurationDataEntry -Node "localhost" `
        -Key "ServerNumber" `
        -Value "0" `
        -Description "Default Value Used to Ensure a Configuration Data File is Generated"

    if ($ConnectionMode -eq 'Credential')
    {
        # Add the GlobalAdminAccount to the Credentials List
        Save-Credentials -UserName "globaladmin"
    }
    else
    {
        Save-Credentials -UserName "certificatepassword"
    }

    $ResourcesPath = Join-Path -Path $PSScriptRoot `
        -ChildPath "..\DSCResources\" `
        -Resolve
    $AllResources = Get-ChildItem $ResourcesPath -Recurse | Where-Object { $_.Name -like 'MSFT_*.psm1' }

    $i = 1
    $ResourcesToExport = @()
    foreach ($ResourceModule in $AllResources)
    {
        try
        {
            $resourceName = $ResourceModule.Name.Split('.')[0].Replace('MSFT_', '')
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
            if (($null -ne $ComponentsToExtract -and
                ($ComponentsToExtract -contains $resourceName -or $ComponentsToExtract -contains ("chck" + $resourceName))) -or
                $AllComponents -or ($null -ne $Workloads -and $Workloads -contains $currentWorkload) -or `
                ($null -eq $ComponentsToExtract -and $null -eq $Workloads) -and `
                ($ComponentsToExtractSpecified -or -not $ComponentsToSkip.Contains($resourceName)))
            {
                $ResourcesToExport += $ResourceModule
            }
        }
        catch
        {
            New-M365DSCLogEntry -Error $_ -Message $ResourceModule.Name -Source "[M365DSCReverse]$($ResourceModule.Name)"
        }
    }

    $platformSkipsNotified = @()
    foreach ($resource in $ResourcesToExport)
    {
        try
        {
            $shouldSkipBecauseOfFailedPlatforms = $false
            $usedPlatforms = Get-ResourcePlatformUsage -Resource $resourceName -ResourceModuleFilePath $ResourceModule.FullName

            foreach($platform in $usedPlatforms)
            {
                # we will skip PnP if there was a problem connecting to a specific site
                # it could be a permissions issue
                # if it was a problem with connecting to the admin site, then we know that all else will fail as well so no need to continue
                if($platform -eq 'PnP' -and $null -ne $Global:SPOAdminUrl -and $Global:SPOConnectionUrl -ne $Global:SPOAdminUrl)
                {
                    continue
                }
                $isAvailable = Check-PlatformAvailability -Platform $platform
                $shouldSkip = $shouldSkip -or !$isAvailable

                if(!$isAvailable -and !$platformSkipsNotified.Contains($platform))
                {
                    Write-Error "The [$platform] connection has failed and all of the related resources will be skipped to avoid unnecessary errors."
                    $platformSkipsNotified += $platform
                }
            }

            if($shouldSkipBecauseOfFailedPlatforms)
            {
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
            if ($GlobalAdminExists-and -not [System.String]::IsNullOrEmpty($GlobalAdminAccount))
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
                    if($parseErrors.Length -gt 0)
                    {
                        Write-Error "The [$resourceName] resource encountered an error and will not be available in the extracted data"
                        Write-Verbose "The [$resourceName] had parse errors"
                        Write-Verbose "#######PARSE ERRORS START####################"
                        foreach($parseError in $parseErrors)
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

                if($psParseErrorOccurred)
                {
                    $exportString = ""
                }
            }
            $DSCContent += $exportString
            $exportString = $null
        }
        catch
        {
            $ex = $_.Exception
            $isMissingGrantError = $false

            while($null -ne $ex -and $ex.GetType().FullName -ne "Microsoft.IdentityModel.Clients.ActiveDirectory.AdalServiceException")
            {
                $ex = $ex.InnerException
            }

            if($null -ne $ex -and $ex.ErrorCode -eq "invalid_grant" -and $null -ne $ex.ServiceErrorCodes -and $ex.ServiceErrorCodes.Contains("65001"))
            {
                $isMissingGrantError = $true
            }

            if($isMissingGrantError -and $currentWorkload -eq 'PP')
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
            $WarningPreference = $DefaultWarningPreference;
            $VerbosePreference = $DefaultVerbosePreference;        
		}
    }

    # Close the Node and Configuration declarations
    $DSCContent += "    }`r`n"
    $DSCContent += "}`r`n"

    if ($ConnectionMode -eq 'Credential')
    {
        #region Add the Prompt for Required Credentials at the top of the Configuration
        $credsContent = ""
        foreach ($credential in $Global:CredsRepo)
        {
            if (!$credential.ToLower().StartsWith("builtin"))
            {
                if (!$AzureAutomation)
                {
                    $credsContent += "        " + (Resolve-Credentials $credential) + " = Get-Credential -Message `"Global Admin credentials`"`r`n"
                }
                else
                {
                    $resolvedName = (Resolve-Credentials $credential)
                    $credsContent += "    " + $resolvedName + " = Get-AutomationPSCredential -Name " + ($resolvedName.Replace("$", "")) + "`r`n"
                }
            }
        }
        $credsContent += "`r`n"
        $startPosition = $DSCContent.IndexOf("<# Credentials #>") + 19
        $DSCContent = $DSCContent.Insert($startPosition, $credsContent)
        $DSCContent += "$ConfigurationName -ConfigurationData .\ConfigurationData.psd1 -GlobalAdminAccount `$GlobalAdminAccount"
        #endregion
    }
    else
    {
        if (-not [System.String]::IsNullOrEmpty($CertificatePassword))
        {
            $certCreds =$Global:CredsRepo[0]
            $credsContent = ""
            $credsContent += "        " + (Resolve-Credentials $certCreds) + " = Get-Credential -Message `"Certificate Password`""
            $credsContent += "`r`n"
            $startPosition = $DSCContent.IndexOf("<# Credentials #>") + 19
            $DSCContent = $DSCContent.Insert($startPosition, $credsContent)
            $DSCContent += "$ConfigurationName -ConfigurationData .\ConfigurationData.psd1 -CertificatePassword `$CertificatePassword"
        }
        else
        {
            $DSCContent += "$ConfigurationName -ConfigurationData .\ConfigurationData.psd1"
        }
    }

    #region Benchmarks
    $M365DSCExportEndTime = [System.DateTime]::Now
    $timeTaken = New-Timespan -Start ($M365DSCExportStartTime.ToString()) `
        -End ($M365DSCExportEndTime.ToString())
    Write-Host "$($Global:M365DSCEmojiHourglass) Export took {" -NoNewLine
    Write-Host "$($timeTaken.TotalSeconds) seconds" -NoNewLine -ForegroundColor Cyan
    Write-Host "}"
    #endregion

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

    if (-not [System.String]::IsNullOrEmpty($FileName))
    {
        $outputDSCFile = $OutputDSCPath + $FileName
    }
    else
    {
        $outputDSCFile = $OutputDSCPath + "M365TenantConfig.ps1"
    }
    
     # this is to avoid problems with unicode charachters when executing the generated ps1 file
    $Utf8BomEncoding = New-Object System.Text.UTF8Encoding $True
    [System.IO.File]::WriteAllText($outputDSCFile, $DSCContent, $Utf8BomEncoding)

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
    $matches = [Regex]::Matches($fileContent, '-Platform\s+(?<platform>\w+)', [ System.Text.RegularExpressions.RegexOptions]::IgnoreCase);

    $platforms = @()
    foreach($match in $matches)
    {
        $platform = $match.Groups["platform"].Value
        if($platforms.Contains($platform))
        {
            continue
        }
        $platforms += $platform
    }

    return $platforms
}
