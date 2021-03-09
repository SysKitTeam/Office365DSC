
#region Session Objects
$Global:SessionSecurityCompliance = $null
#endregion

#region Extraction Modes
$Global:DefaultComponents = @("SPOApp", "SPOSiteDesign")
$Global:FullComponents = @("AADMSGroup", "AADServicePrincipal", "EXOMailboxSettings", "EXOManagementRole", "O365Group", "O365User", `
        "PlannerPlan", "PlannerBucket", "PlannerTask", "PPPowerAppsEnvironment", `
        "SPOSiteAuditSettings", "SPOSiteGroup", "SPOSite", "SPOUserProfileProperty", "SPOPropertyBag", "TeamsTeam", "TeamsChannel", `
        "TeamsUser", "TeamsChannelTab")
#endregion

function Format-EXOParams
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [System.Collections.Hashtable]
        $InputEXOParams,

        [Parameter()]
        [ValidateSet('New', 'Set')]
        [System.String]
        $Operation
    )
    $EXOParams = $InputEXOParams
    $EXOParams.Remove("GlobalAdminAccount") | Out-Null
    $EXOParams.Remove("Ensure") | Out-Null
    $EXOParams.Remove("Verbose") | Out-Null
    if ('New' -eq $Operation)
    {
        $EXOParams += @{
            Name = $EXOParams.Identity
        }
        $EXOParams.Remove("Identity") | Out-Null
        $EXOParams.Remove("MakeDefault") | Out-Null
        return $EXOParams
    }
    if ('Set' -eq $Operation)
    {
        $EXOParams.Remove("Enabled") | Out-Null
        return $EXOParams
    }
}

function Get-TimeZoneNameFromID
{
    [CmdletBinding()]
    [OutputType([String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ID
    )

    $TimezoneObject = $Timezones | Where-Object -FilterScript { $_.ID -eq $ID }

    if ($null -eq $TimezoneObject)
    {
        throw "The specified Timzone with ID {$($ID)} is not valid"
    }
    return $TimezoneObject.EnglishName
}
function Get-TimeZoneIDFromName
{
    [CmdletBinding()]
    [OutputType([String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Name
    )

    $TimezoneObject = $Timezones | Where-Object -FilterScript { $_.EnglishName -eq $Name }

    if ($null -eq $TimezoneObject)
    {
        throw "The specified Timzone {$($Name)} is not valid"
    }
    return $TimezoneObject.ID
}

function Get-TeamByGroupID
{
    [CmdletBinding()]
    [OutputType([Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $GroupId
    )

    $team = Get-Team -GroupId $GroupId
    if ($null -eq $team)
    {
        return $false
    }
    return $true
}
function Get-TeamByName
{
    [CmdletBinding()]
    [OutputType([Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $TeamName
    )

    $loopCounter = 0
    do
    {
        $team = Get-Team -DisplayName $TeamName
        if ($null -eq $team)
        {
            Start-Sleep 5
        }
        $loopCounter += 1
        if ($loopCounter -gt 5)
        {
            break
        }
    } while ($null -eq $team)

    if ($null -eq $team)
    {
        throw "Team with Name $TeamName doesn't exist in tenant"
    }
    return $team
}


function Reset-AllTeamsCached
{
    [CmdletBinding()]
    param
    (
    )

    $Global:O365TeamsCached = $null
}


function Get-TeamEnabledOffice365Groups
{
    try
    {
        $allTeams = New-Object Collections.Generic.List[Microsoft.TeamsCmdlets.PowerShell.Custom.Model.Team]
        $endpoint = Get-AzureEnvironmentEndpoint -AzureCloudEnvironmentName $Global:appIdentityParams.AzureCloudEnvironmentName -EndpointName "MsGraphEndpointResourceId"
        $accessToken = Get-AppIdentityAccessToken $endpoint

        $clientFactory = [Microsoft.TeamsCmdlets.PowerShell.Custom.Utils.HttpClientFactory]::new()
        $httpClient = $clientFactory.Create("Bearer $accessToken", "Get-TeamTraceCustom")

        $requestUri = [Uri]::new("$endpoint/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&`$select=id,resourceProvisioningOptions,displayName,description,visibility,mailnickname,classification")
        Invoke-WithTransientErrorExponentialRetry -ScriptBlock {
            $accessToken = Get-AppIdentityAccessToken $endpoint
            $httpClient.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::Parse("Bearer $accessToken");
            $allTeams.AddRange([Microsoft.TeamsCmdlets.PowerShell.Custom.Utils.HttpUtilities].GetMethod("GetAll").MakeGenericMethod([Microsoft.TeamsCmdlets.PowerShell.Custom.Model.Team]).Invoke($null, @($httpClient, $requestUri)))
            Write-Verbose "Retrieved all teams"
        }

        $allTeams = $allTeams | Where-Object {
            $_.ResourceProvisioningOptions.Contains("Team")
        }
    }
    finally
    {
        if ($null -ne $httpClient)
        {
            $httpClient.Dispose()
        }
    }

    return $allTeams
}



function Get-AllTeamsCached
{
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [Switch]
        $ForceRefresh
    )

    if ($Global:O365TeamsCached -and !$ForceRefresh)
    {
        return $Global:O365TeamsCached
    }

    $allTeamSettings = New-Object Collections.Generic.List[Microsoft.TeamsCmdlets.PowerShell.Custom.Model.TeamSettings]

    $endpoint = Get-AzureEnvironmentEndpoint -AzureCloudEnvironmentName $Global:appIdentityParams.AzureCloudEnvironmentName -EndpointName "MsGraphEndpointResourceId"

    # The Get-Team cmdlet was not really written with throttling in mind
    # internally, they get the list of teams and then in parallel go to the /teams/{id} endpoint
    # this is actually the only way to get the team details, but when running in parallel without any limits
    # throttling is bound to come up and it is NOT handled at all
    # Get-Team
    $clientFactory = [Microsoft.TeamsCmdlets.PowerShell.Custom.Utils.HttpClientFactory]::new()
    $accessToken = Get-AppIdentityAccessToken $endpoint
    $allTeams = Get-TeamEnabledOffice365Groups

    $allTeams | ForEach-Object {
        try
        {
            $singleTeamClient = $clientFactory.Create("Bearer $accessToken", "Get-TeamTraceCustom")
            $teamToRetrieve = $_
            Invoke-WithTransientErrorExponentialRetry -ScriptBlock {
                $accessToken = Get-AppIdentityAccessToken $endpoint
                $singleTeamClient.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::Parse("Bearer $accessToken")
                $groupId = $teamToRetrieve.GroupId
                $singleTeamRequestUri = [Uri]::new("$endpoint/v1.0/teams/$groupId")
                Write-Verbose "retrieving from $singleTeamRequestUri"
                [Type[]]$types = @([System.Net.Http.HttpClient], [Uri])
                $team = [Microsoft.TeamsCmdlets.PowerShell.Custom.Utils.HttpUtilities].GetMethod("Get", $types).MakeGenericMethod([Microsoft.TeamsCmdlets.PowerShell.Custom.Model.Team]).Invoke($null, @($singleTeamClient, $singleTeamRequestUri))
                $team.DisplayName = $_.DisplayName
                $team.Description = $_.Description
                $team.Visibility = $_.Visibility
                $team.MailNickName = $_.MailNickName
                $team.Classification = $_.Classification

                $allTeamSettings.Add([Microsoft.TeamsCmdlets.PowerShell.Custom.Model.TeamSettings]::new($team))
            }
        }
        catch
        {
            $missingTeam = ($null -ne $_.Exception -and $_.Exception.ErrorCode -eq 404) -or ($null -ne $_.Exception -and $null -ne $_.Exception.InnerException -and $_.Exception.InnerException.ErrorCode -eq 404);

            # write the output only if the teams is not missing
            # if it's missing then it was probably deleted or something like that
            if (!$missingTeam)
            {
                # we write it with verbose because if one team retrieval fails, they probably all will. No teams is hard to miss in the output data so the error should be evident
                Write-Verbose $_
            }

        }
        finally
        {
            if ($null -ne $singleTeamClient)
            {
                $singleTeamClient.Dispose()
            }
        }
    }

    $Global:O365TeamsCached = $allTeamSettings
    return $Global:O365TeamsCached
}


function Invoke-WithTransientErrorExponentialRetry
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ScriptBlock]
        $ScriptBlock
    )

    if ($null -eq $Global:O365DSCBackoffRandomizer)
    {
        $Global:O365DSCBackoffRandomizer = [System.Random]::new()
    }

    $retryCount = 10
    $backoffPeriodMiliseconds = 500
    do
    {
        try
        {
            Write-Verbose "Executing script block, retryCount: $retryCount"
            Invoke-Command $ScriptBlock
            break
        }
        catch
        {
            if ($null -eq $_.Exception -or $null -eq $_.Exception.InnerException -or ($_.Exception.InnerException.ErrorCode -ne 429 -and $_.Exception.InnerException.ErrorCode -ne 503) -or $retryCount -eq 0)
            {
                throw
            }

            Write-Verbose "The request has been throttled, will retry, retryCount: $retryCount"
        }
        Start-Sleep -Milliseconds $backoffPeriodMiliseconds
        $retryCount = $retryCount - 1
        $backoffPeriodMiliseconds = $backoffPeriodMiliseconds * 2 + $Global:O365DSCBackoffRandomizer.Next($backoffPeriodMiliseconds / 10)
    }
    while ($retryCount -gt 0)
}


function Convert-M365DscHashtableToString
{
    param
    (
        [Parameter()]
        [System.Collections.Hashtable]
        $Hashtable
    )
    $values = @()
    foreach ($pair in $Hashtable.GetEnumerator())
    {
        try
        {
            if ($pair.Value -is [System.Array])
            {
                $str = "$($pair.Key)=($($pair.Value -join ","))"
            }
            elseif ($pair.Value -is [System.Collections.Hashtable])
            {
                $str = "$($pair.Key)={$(Convert-M365DscHashtableToString -Hashtable $pair.Value)}"
            }
            else
            {
                if ($null -eq $pair.Value)
                {
                    $str = "$($pair.Key)=`$null"
                }
                else
                {
                    $str = "$($pair.Key)=$($pair.Value)"
                }
            }
            $values += $str
        }
        catch
        {
            Write-Warning "There was an error converting the Hashtable to a string: $_"
        }
    }

    [array]::Sort($values)
    return ($values -join "; ")
}

function New-EXOAntiPhishPolicy
{
    param (
        [Parameter()]
        [System.Collections.Hashtable]
        $AntiPhishPolicyParams
    )
    try
    {
        $VerbosePreference = 'Continue'
        $BuiltParams = (Format-EXOParams -InputEXOParams $AntiPhishPolicyParams -Operation 'New' )
        Write-Verbose -Message "Creating New AntiPhishPolicy $($BuiltParams.Name) with values: $(Convert-M365DscHashtableToString -Hashtable $BuiltParams)"
        New-AntiPhishPolicy @BuiltParams
        $VerbosePreference = 'SilentlyContinue'
    }
    catch
    {
        Close-SessionsAndReturnError -ExceptionMessage $_.Exception
    }
}

function New-EXOSafeAttachmentRule
{
    param (
        [Parameter()]
        [System.Collections.Hashtable]
        $SafeAttachmentRuleParams
    )
    try
    {
        $VerbosePreference = 'Continue'
        $BuiltParams = (Format-EXOParams -InputEXOParams $SafeAttachmentRuleParams -Operation 'New' )
        Write-Verbose -Message "Creating New SafeAttachmentRule $($BuiltParams.Name) with values: $(Convert-M365DscHashtableToString -Hashtable $BuiltParams)"
        New-SafeAttachmentRule @BuiltParams -Confirm:$false
        $VerbosePreference = 'SilentlyContinue'
    }
    catch
    {
        Close-SessionsAndReturnError -ExceptionMessage $_.Exception
    }
}

function New-EXOSafeLinksRule
{
    param (
        [Parameter()]
        [System.Collections.Hashtable]
        $SafeLinksRuleParams
    )
    try
    {
        $VerbosePreference = 'Continue'
        $BuiltParams = (Format-EXOParams -InputEXOParams $SafeLinksRuleParams -Operation 'New' )
        Write-Verbose -Message "Creating New SafeLinksRule $($BuiltParams.Name) with values: $(Convert-M365DscHashtableToString -Hashtable $BuiltParams)"
        New-SafeLinksRule @BuiltParams -Confirm:$false
        $VerbosePreference = 'SilentlyContinue'
    }
    catch
    {
        Close-SessionsAndReturnError -ExceptionMessage $_.Exception
    }
}

function Set-EXOAntiPhishPolicy
{
    param (
        [Parameter()]
        [System.Collections.Hashtable]
        $AntiPhishPolicyParams
    )
    try
    {
        $VerbosePreference = 'Continue'
        $BuiltParams = (Format-EXOParams -InputEXOParams $AntiPhishPolicyParams -Operation 'Set' )
        if ($BuiltParams.keys -gt 1)
        {
            Write-Verbose -Message "Setting AntiPhishPolicy $($BuiltParams.Identity) with values: $(Convert-M365DscHashtableToString -Hashtable $BuiltParams)"
            Set-AntiPhishPolicy @BuiltParams -Confirm:$false
            $VerbosePreference = 'SilentlyContinue'
        }
        else
        {
            Write-Verbose -Message "No more values to Set on AntiPhishPolicy $($BuiltParams.Identity) using supplied values: $(Convert-M365DscHashtableToString -Hashtable $BuiltParams)"
            $VerbosePreference = 'SilentlyContinue'
        }
    }
    catch
    {
        Close-SessionsAndReturnError -ExceptionMessage $_.Exception
    }
}

function Confirm-ImportedCmdletIsAvailable
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $CmdletName
    )
    try
    {
        $CmdletIsAvailable = (Get-Command -Name $CmdletName -ErrorAction SilentlyContinue)
        if ($CmdletIsAvailable)
        {
            return $true
        }
        else
        {
            return $false
        }
    }
    catch
    {
        return $false
    }
}

function Set-EXOSafeAttachmentRule
{
    param (
        [Parameter()]
        [System.Collections.Hashtable]
        $SafeAttachmentRuleParams
    )
    try
    {
        $VerbosePreference = 'Continue'
        $BuiltParams = (Format-EXOParams -InputEXOParams $SafeAttachmentRuleParams -Operation 'Set' )
        if ($BuiltParams.keys -gt 1)
        {
            Write-Verbose -Message "Setting SafeAttachmentRule $($BuiltParams.Identity) with values: $(Convert-M365DscHashtableToString -Hashtable $BuiltParams)"
            Set-SafeAttachmentRule @BuiltParams -Confirm:$false
            $VerbosePreference = 'SilentlyContinue'
        }
        else
        {
            Write-Verbose -Message "No more values to Set on SafeAttachmentRule $($BuiltParams.Identity) using supplied values: $(Convert-M365DscHashtableToString -Hashtable $BuiltParams)"
            $VerbosePreference = 'SilentlyContinue'
        }
    }
    catch
    {
        Close-SessionsAndReturnError -ExceptionMessage $_.Exception
    }
}

function Set-EXOSafeLinksRule
{
    param (
        [Parameter()]
        [System.Collections.Hashtable]
        $SafeLinksRuleParams
    )
    try
    {
        $VerbosePreference = 'Continue'
        $BuiltParams = (Format-EXOParams -InputEXOParams $SafeLinksRuleParams -Operation 'Set' )
        if ($BuiltParams.keys -gt 1)
        {
            Write-Verbose -Message "Setting SafeLinksRule $($BuiltParams.Identity) with values: $(Convert-M365DscHashtableToString -Hashtable $BuiltParams)"
            Set-SafeLinksRule @BuiltParams -Confirm:$false
            $VerbosePreference = 'SilentlyContinue'
        }
        else
        {
            Write-Verbose -Message "No more values to Set on SafeLinksRule $($BuiltParams.Identity) using supplied values: $(Convert-M365DscHashtableToString -Hashtable $BuiltParams)"
            $VerbosePreference = 'SilentlyContinue'
        }
    }
    catch
    {
        Close-SessionsAndReturnError -ExceptionMessage $_.Exception
    }
}

function Compare-PSCustomObjectArrays
{
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Object[]]
        $DesiredValues,

        [Parameter(Mandatory = $true)]
        [System.Object[]]
        $CurrentValues
    )

    $DriftedProperties = @()
    foreach ($DesiredEntry in $DesiredValues)
    {
        $Properties = $DesiredEntry.PSObject.Properties
        $KeyProperty = $Properties.Name[0]

        $EquivalentEntryInCurrent = $CurrentValues | Where-Object -FilterScript { $_.$KeyProperty -eq $DesiredEntry.$KeyProperty }
        if ($null -eq $EquivalentEntryInCurrent)
        {
            $result = @{
                Property     = $DesiredEntry
                PropertyName = $KeyProperty
                Desired      = $DesiredEntry.$KeyProperty
                Current      = $null
            }
            $DriftedProperties += $DesiredEntry
        }
        else
        {
            foreach ($property in $Properties)
            {
                $propertyName = $property.Name

                if ($DesiredEntry.$PropertyName -ne $EquivalentEntryInCurrent.$PropertyName)
                {
                    $result = @{
                        Property     = $DesiredEntry
                        PropertyName = $PropertyName
                        Desired      = $DesiredEntry.$PropertyName
                        Current      = $EquivalentEntryInCurrent.$PropertyName
                    }
                    $DriftedProperties += $result
                }
            }
        }
    }

    return $DriftedProperties
}

function Test-M365DSCParameterState
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true, Position = 1)]
        [HashTable]
        $CurrentValues,

        [Parameter(Mandatory = $true, Position = 2)]
        [Object]
        $DesiredValues,

        [Parameter(Position = 3)]
        [Array]
        $ValuesToCheck,

        [Parameter(Position = 4)]
        [System.String]
        $Source = 'Generic'
    )
    #region Telemetry
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", "$Source")
    $data.Add("Method", "Test-TargetResource")
    #endregion
    $returnValue = $true

    $DriftedParameters = @{ }

    if (($DesiredValues.GetType().Name -ne "HashTable") `
            -and ($DesiredValues.GetType().Name -ne "CimInstance") `
            -and ($DesiredValues.GetType().Name -ne "PSBoundParametersDictionary"))
    {
        throw ("Property 'DesiredValues' in Test-M365DSCParameterState must be either a " + `
                "Hashtable or CimInstance. Type detected was $($DesiredValues.GetType().Name)")
    }

    if (($DesiredValues.GetType().Name -eq "CimInstance") -and ($null -eq $ValuesToCheck))
    {
        throw ("If 'DesiredValues' is a CimInstance then property 'ValuesToCheck' must contain " + `
                "a value")
    }

    if (($null -eq $ValuesToCheck) -or ($ValuesToCheck.Count -lt 1))
    {
        $KeyList = $DesiredValues.Keys
    }
    else
    {
        $KeyList = $ValuesToCheck
    }

    $KeyList | ForEach-Object -Process {
        if (($_ -ne "Verbose") -and ($_ -ne "InstallAccount"))
        {
            if (($CurrentValues.ContainsKey($_) -eq $false) `
                    -or ($CurrentValues.$_ -ne $DesiredValues.$_) `
                    -or (($DesiredValues.ContainsKey($_) -eq $true) -and ($null -ne $DesiredValues.$_ -and $DesiredValues.$_.GetType().IsArray)))
            {
                if ($DesiredValues.GetType().Name -eq "HashTable" -or `
                        $DesiredValues.GetType().Name -eq "PSBoundParametersDictionary")
                {
                    $CheckDesiredValue = $DesiredValues.ContainsKey($_)
                }
                else
                {
                    $CheckDesiredValue = Test-M365DSCObjectHasProperty -Object $DesiredValues -PropertyName $_
                }

                if ($CheckDesiredValue)
                {
                    $desiredType = $DesiredValues.$_.GetType()
                    $fieldName = $_
                    if ($desiredType.IsArray -eq $true)
                    {
                        if (($CurrentValues.ContainsKey($fieldName) -eq $false) `
                                -or ($null -eq $CurrentValues.$fieldName))
                        {
                            Write-Verbose -Message ("Expected to find an array value for " + `
                                    "property $fieldName in the current " + `
                                    "values, but it was either not present or " + `
                                    "was null. This has caused the test method " + `
                                    "to return false.")
                            $DriftedParameters.Add($fieldName, '')
                            $returnValue = $false
                        }
                        elseif ($desiredType.Name -eq 'ciminstance[]')
                        {
                            Write-Verbose "The current property {$_} is a CimInstance[]"
                            $AllDesiredValuesAsArray = @()
                            foreach ($item in $DesiredValues.$_)
                            {
                                $currentEntry = @{ }
                                foreach ($prop in $item.CIMInstanceProperties)
                                {
                                    $value = $prop.Value
                                    if ([System.String]::IsNullOrEmpty($value))
                                    {
                                        $value = $null
                                    }
                                    $currentEntry.Add($prop.Name, $value)
                                }
                                $AllDesiredValuesAsArray += [PSCustomObject]$currentEntry
                            }

                            $arrayCompare = Compare-PSCustomObjectArrays -CurrentValues $CurrentValues.$fieldName `
                                -DesiredValues $AllDesiredValuesAsArray
                            if ($null -ne $arrayCompare)
                            {
                                foreach ($item in $arrayCompare)
                                {
                                    $EventValue = "<CurrentValue>[$($item.PropertyName)]$($item.CurrentValue)</CurrentValue>"
                                    $EventValue += "<DesiredValue>[$($item.PropertyName)]$($item.DesiredValue)</DesiredValue>"
                                    $DriftedParameters.Add($fieldName, $EventValue)
                                }
                                $returnValue = $false
                            }
                        }
                        else
                        {
                            $arrayCompare = Compare-Object -ReferenceObject $CurrentValues.$fieldName `
                                -DifferenceObject $DesiredValues.$fieldName
                            if ($null -ne $arrayCompare -and
                                -not [System.String]::IsNullOrEmpty($arrayCompare.InputObject))
                            {
                                Write-Verbose -Message ("Found an array for property $fieldName " + `
                                        "in the current values, but this array " + `
                                        "does not match the desired state. " + `
                                        "Details of the changes are below.")
                                $arrayCompare | ForEach-Object -Process {
                                    Write-Verbose -Message "$($_.InputObject) - $($_.SideIndicator)"
                                }

                                $EventValue = "<CurrentValue>$($CurrentValues.$fieldName)</CurrentValue>"
                                $EventValue += "<DesiredValue>$($DesiredValues.$fieldName)</DesiredValue>"
                                $DriftedParameters.Add($fieldName, $EventValue)
                                $returnValue = $false
                            }
                        }
                    }
                    else
                    {
                        switch ($desiredType.Name)
                        {
                            "String"
                            {
                                if ([string]::IsNullOrEmpty($CurrentValues.$fieldName) `
                                        -and [string]::IsNullOrEmpty($DesiredValues.$fieldName))
                                {
                                }
                                else
                                {
                                    Write-Verbose -Message ("String value for property " + `
                                            "$fieldName does not match. " + `
                                            "Current state is " + `
                                            "'$($CurrentValues.$fieldName)' " + `
                                            "and desired state is " + `
                                            "'$($DesiredValues.$fieldName)'")
                                    $EventValue = "<CurrentValue>$($CurrentValues.$fieldName)</CurrentValue>"
                                    $EventValue += "<DesiredValue>$($DesiredValues.$fieldName)</DesiredValue>"
                                    $DriftedParameters.Add($fieldName, $EventValue)
                                    $returnValue = $false
                                }
                            }
                            "Int32"
                            {
                                if (($DesiredValues.$fieldName -eq 0) `
                                        -and ($null -eq $CurrentValues.$fieldName))
                                {
                                }
                                else
                                {
                                    Write-Verbose -Message ("Int32 value for property " + `
                                            "$fieldName does not match. " + `
                                            "Current state is " + `
                                            "'$($CurrentValues.$fieldName)' " + `
                                            "and desired state is " + `
                                            "'$($DesiredValues.$fieldName)'")
                                    $EventValue = "<CurrentValue>$($CurrentValues.$fieldName)</CurrentValue>"
                                    $EventValue += "<DesiredValue>$($DesiredValues.$fieldName)</DesiredValue>"
                                    $DriftedParameters.Add($fieldName, $EventValue)
                                    $returnValue = $false
                                }
                            }
                            "Int16"
                            {
                                if (($DesiredValues.$fieldName -eq 0) `
                                        -and ($null -eq $CurrentValues.$fieldName))
                                {
                                }
                                else
                                {
                                    Write-Verbose -Message ("Int16 value for property " + `
                                            "$fieldName does not match. " + `
                                            "Current state is " + `
                                            "'$($CurrentValues.$fieldName)' " + `
                                            "and desired state is " + `
                                            "'$($DesiredValues.$fieldName)'")
                                    $EventValue = "<CurrentValue>$($CurrentValues.$fieldName)</CurrentValue>"
                                    $EventValue += "<DesiredValue>$($DesiredValues.$fieldName)</DesiredValue>"
                                    $DriftedParameters.Add($fieldName, $EventValue)
                                    $returnValue = $false
                                }
                            }
                            "Boolean"
                            {
                                if ($CurrentValues.$fieldName -ne $DesiredValues.$fieldName)
                                {
                                    Write-Verbose -Message ("Boolean value for property " + `
                                            "$fieldName does not match. " + `
                                            "Current state is " + `
                                            "'$($CurrentValues.$fieldName)' " + `
                                            "and desired state is " + `
                                            "'$($DesiredValues.$fieldName)'")
                                    $EventValue = "<CurrentValue>$($CurrentValues.$fieldName)</CurrentValue>"
                                    $EventValue += "<DesiredValue>$($DesiredValues.$fieldName)</DesiredValue>"
                                    $DriftedParameters.Add($fieldName, $EventValue)
                                    $returnValue = $false
                                }
                            }
                            "Single"
                            {
                                if (($DesiredValues.$fieldName -eq 0) `
                                        -and ($null -eq $CurrentValues.$fieldName))
                                {
                                }
                                else
                                {
                                    Write-Verbose -Message ("Single value for property " + `
                                            "$fieldName does not match. " + `
                                            "Current state is " + `
                                            "'$($CurrentValues.$fieldName)' " + `
                                            "and desired state is " + `
                                            "'$($DesiredValues.$fieldName)'")
                                    $EventValue = "<CurrentValue>$($CurrentValues.$fieldName)</CurrentValue>"
                                    $EventValue += "<DesiredValue>$($DesiredValues.$fieldName)</DesiredValue>"
                                    $DriftedParameters.Add($fieldName, $EventValue)
                                    $returnValue = $false
                                }
                            }
                            "Hashtable"
                            {
                                Write-Verbose -Message "The current property {$fieldName} is a Hashtable"
                                $AllDesiredValuesAsArray = @()
                                foreach ($item in $DesiredValues.$fieldName)
                                {
                                    $currentEntry = @{ }
                                    foreach ($key in $item.Keys)
                                    {
                                        $value = $item.$key
                                        if ([System.String]::IsNullOrEmpty($value))
                                        {
                                            $value = $null
                                        }
                                        $currentEntry.Add($key, $value)
                                    }
                                    $AllDesiredValuesAsArray += [PSCustomObject]$currentEntry
                                }

                                if ($null -ne $DesiredValues.$fieldName -and $null -eq $CurrentValues.$fieldName)
                                {
                                    $returnValue = $false
                                }
                                else
                                {
                                    $AllCurrentValuesAsArray = @()
                                    foreach ($item in $CurrentValues.$fieldName)
                                    {
                                        $currentEntry = @{ }
                                        foreach ($key in $item.Keys)
                                        {
                                            $value = $item.$key
                                            if ([System.String]::IsNullOrEmpty($value))
                                            {
                                                $value = $null
                                            }
                                            $currentEntry.Add($key, $value)
                                        }
                                        $AllCurrentValuesAsArray += [PSCustomObject]$currentEntry
                                    }
                                    $arrayCompare = Compare-PSCustomObjectArrays -CurrentValues $AllCurrentValuesAsArray `
                                        -DesiredValues $AllDesiredValuesAsArray
                                    if ($null -ne $arrayCompare)
                                    {
                                        foreach ($item in $arrayCompare)
                                        {
                                            $EventValue = "<CurrentValue>[$($item.PropertyName)]$($item.CurrentValue)</CurrentValue>"
                                            $EventValue += "<DesiredValue>[$($item.PropertyName)]$($item.DesiredValue)</DesiredValue>"
                                            $DriftedParameters.Add($fieldName, $EventValue)
                                        }
                                        $returnValue = $false
                                    }
                                }
                            }
                            default
                            {
                                Write-Verbose -Message ("Unable to compare property $fieldName " + `
                                        "as the type ($($desiredType.Name)) is " + `
                                        "not handled by the " + `
                                        "Test-M365DSCParameterState cmdlet")
                                $EventValue = "<CurrentValue>$($CurrentValues.$fieldName)</CurrentValue>"
                                $EventValue += "<DesiredValue>$($DesiredValues.$fieldName)</DesiredValue>"
                                $DriftedParameters.Add($fieldName, $EventValue)
                                $returnValue = $false
                            }
                        }
                    }
                }
            }
        }
    }

    if ($returnValue -eq $false)
    {
        $EventMessage = "<M365DSCEvent>`r`n"
        $EventMessage += "    <ConfigurationDrift Source=`"$Source`">`r`n"

        $EventMessage += "        <ParametersNotInDesiredState>`r`n"
        $driftedValue = ''
        foreach ($key in $DriftedParameters.Keys)
        {
            Write-Verbose -Message "Detected Drifted Parameter [$Source]$key"
            #region Telemetry
            $driftedData = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
            $driftedData.Add("Event", "DriftedParameter")
            $driftedData.Add("Parameter", "[$Source]$key")
            Add-M365DSCTelemetryEvent -Type "DriftInfo" -Data $driftedData
            #endregion
            $EventMessage += "            <Param Name=`"$key`">" + $DriftedParameters.$key + "</Param>`r`n"
        }
        #region Telemetry
        $data.Add("Event", "ConfigurationDrift")
        #endregion
        $EventMessage += "        </ParametersNotInDesiredState>`r`n"
        $EventMessage += "    </ConfigurationDrift>`r`n"
        $EventMessage += "    <DesiredValues>`r`n"
        foreach ($Key in $DesiredValues.Keys)
        {
            $Value = $DesiredValues.$Key
            if ([System.String]::IsNullOrEmpty($Value))
            {
                $Value = "`$null"
            }
            $EventMessage += "        <Param Name =`"$key`">$Value</Param>`r`n"
        }
        $EventMessage += "    </DesiredValues>`r`n"
        $EventMessage += "</M365DSCEvent>"

        Add-M365DSCEvent -Message $EventMessage -EntryType 'Warning' `
            -EventID 1 -Source $Source
    }
    #region Telemetry
    Add-M365DSCTelemetryEvent -Data $data
    #endregion
    return $returnValue
}

<# This is the main Microsoft365DSC.Reverse function that extracts the DSC configuration from an existing
   Office 365 Tenant. #>
function Export-M365DSCConfiguration
{
    [CmdletBinding()]
    param(
        [Parameter()]
        [Switch]
        $Quiet,

        [Parameter()]
        [System.String]
        $Path,

        [Parameter()]
        [System.String]
        $FileName,

        [Parameter()]
        [System.String]
        $ConfigurationName,

        [Parameter()]
        [System.String[]]
        $ComponentsToExtract,

        [Parameter()]
        [Switch]
        $AllComponents,

        [Parameter()]
        [ValidateSet('AAD', 'SPO', 'EXO', 'INTUNE', 'SC', 'OD', 'O365', 'PLANNER', 'PP', 'TEAMS')]
        [System.String[]]
        $Workloads,

        [Parameter()]
        [ValidateRange(1, 100)]
        $MaxProcesses,

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
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword,

        [Parameter()]
        [System.String]
        $CertificatePath
    )
    #region Telemetry
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Event", "Extraction")
    $data.Add("Quiet", $Quiet)
    $data.Add("Path", [System.String]::IsNullOrEmpty($Path))
    $data.Add("FileName", $null -ne [System.String]::IsNullOrEmpty($FileName))
    $data.Add("ComponentsToExtract", $null -ne $ComponentsToExtract)
    $data.Add("Workloads", $null -ne $Workloads)
    $data.Add("MaxProcesses", $null -ne $MaxProcesses)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    if ($null -eq $MaxProcesses)
    {
        $MaxProcesses = 16
    }

    if (-not $Quiet)
    {
        Show-M365DSCGUI -Path $Path -FileName $FileName `
            -GenerateInfo $GenerateInfo
    }
    else
    {
        if ($null -ne $Workloads)
        {
            Start-M365DSCConfigurationExtract -GlobalAdminAccount $GlobalAdminAccount `
                -Workloads $Workloads `
                -Mode $Mode `
                -Path $Path -FileName $FileName `
                -MaxProcesses $MaxProcesses `
                -ConfigurationName $ConfigurationName `
                -ApplicationId $ApplicationId `
                -TenantId $TenantId `
                -CertificateThumbprint $CertificateThumbprint `
                -CertificatePath $CertificatePath `
                -CertificatePassword $CertificatePassword `
                -GenerateInfo $GenerateInfo `
                -Quiet
        }
        elseif ($null -ne $ComponentsToExtract)
        {
            Start-M365DSCConfigurationExtract -GlobalAdminAccount $GlobalAdminAccount `
                -ComponentsToExtract $ComponentsToExtract `
                -Path $Path -FileName $FileName `
                -MaxProcesses $MaxProcesses `
                -ConfigurationName $ConfigurationName `
                -ApplicationId $ApplicationId `
                -TenantId $TenantId `
                -CertificateThumbprint $CertificateThumbprint `
                -CertificatePath $CertificatePath `
                -CertificatePassword $CertificatePassword `
                -GenerateInfo $GenerateInfo `
                -Quiet
        }
        elseif ($AllComponents)
        {
            Start-M365DSCConfigurationExtract -GlobalAdminAccount $GlobalAdminAccount `
                -Path $Path -FileName $FileName `
                -AllComponents `
                -MaxProcesses $MaxProcesses `
                -ConfigurationName $ConfigurationName `
                -ApplicationId $ApplicationId `
                -TenantId $TenantId `
                -CertificateThumbprint $CertificateThumbprint `
                -CertificatePath $CertificatePath `
                -CertificatePassword $CertificatePassword `
                -GenerateInfo $GenerateInfo `
                -Quiet
        }
    }
}

function Get-M365DSCTenantDomain
{
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $ApplicationId,

        [Parameter(Mandatory = $true)]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [System.String]
        $CertificatePath
    )

    if (!$CertificatePath)
    {
        $ConnectionMode = New-M365DSCConnection -Platform 'AzureAD' `
            -InboundParameters $PSBoundParameters
        $tenantDetails = Get-AzureADTenantDetail
        $defaultDomain = $tenantDetails.VerifiedDomains | Where-Object -FilterScript { $_.Initial }
        return $defaultDomain.Name
    }
    if ($TenantId.Contains("onmicrosoft"))
    {
        return $TenantId
    }
    else
    {
        throw "TenantID must be in format contoso.onmicrosoft.com"
    }

}

function Get-M365DSCOrganization
{
    param(
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount,

        [Parameter()]
        [System.String]
        $TenantId

    )
    if ($null -ne $GlobalAdminAccount -and $GlobalAdminAccount.UserName.Contains("@"))
    {
        $organization = $GlobalAdminAccount.UserName.Split("@")[1]
        return $organization
    }
    if (-not [System.String]::IsNullOrEmpty($TenantId))
    {
        if ($TenantId.contains("."))
        {
            $organization = $TenantId
            return $organization
        }
        else
        {
            Throw "Tenant ID must be name of tenant not a GUID. Ex contoso.onmicrosoft.com"
        }

    }
}

function New-M365DSCConnection
{
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Azure", "AzureAD", "ExchangeOnline", "Intune", `
                "SecurityComplianceCenter", "PnP", "PowerPlatforms", `
                "MicrosoftTeams", "SkypeForBusiness", "MicrosoftGraph", `
                "MicrosoftGraphBeta")]
        [System.String]
        $Platform,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $InboundParameters,

        [Parameter()]
        [System.String]
        $Url,

        [Parameter()]
        [System.Boolean]
        $SkipModuleReload = $false
    )

    Write-Verbose -Message "Attempting connection to {$Platform} with:"
    Write-Verbose -Message "$($InboundParameters | Out-String)"

    if ($SkipModuleReload -eq $true)
    {
        $Global:CurrentModeIsExport = $true
    }
    else
    {
        $Global:CurrentModeIsExport = $false
    }

    Test-MSCloudLogin -Platform $Platform -CloudCredential $InboundParameters.GlobalAdminAccount -ConnectionUrl $Url
    return "ServicePrincipal"

    # Case both authentication methods are attempted
    if ($null -ne $InboundParameters.GlobalAdminAccount -and `
        (-not [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -or `
                -not [System.String]::IsNullOrEmpty($InboundParameters.CertificateThumbprint)))
    {
        $message = 'Both Authentication methods are attempted'
        Write-Verbose -Message $message
        $data.Add("Event", "Error")
        $data.Add("Exception", $message)
        $errorText = "You can't specify both the GlobalAdminAccount and one of {TenantId, CertificateThumbprint}"
        $data.Add("CustomMessage", $errorText)
        Add-M365DSCTelemetryEvent -Type "Error" -Data $data
        throw $errorText
    }
    # Case no authentication method is specified
    elseif ($null -eq $InboundParameters.GlobalAdminAccount -and `
            [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId) -and `
            [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -and `
            [System.String]::IsNullOrEmpty($InboundParameters.CertificateThumbprint))
    {
        $message = 'No Authentication method was provided'
        Write-Verbose -Message $message
        $data.Add("Event", "Error")
        $data.Add("Exception", $message)
        $errorText = "You must specify either the GlobalAdminAccount or ApplicationId, TenantId and CertificateThumbprint parameters."
        $data.Add("CustomMessage", $errorText)
        Add-M365DSCTelemetryEvent -Type "Error" -Data $data
        throw $errorText
    }
    # Case only GlobalAdminAccount is specified
    elseif ($null -ne $InboundParameters.GlobalAdminAccount -and `
            [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId) -and `
            [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -and `
            [System.String]::IsNullOrEmpty($InboundParameters.CertificateThumbprint))
    {
        Write-Verbose -Message "GlobalAdminAccount was specified. Connecting via User Principal"
        if ([System.String]::IsNullOrEmpty($Url))
        {
            Test-MSCloudLogin -Platform $Platform `
                -CloudCredential $InboundParameters.GlobalAdminAccount `
                -SkipModuleReload $Global:CurrentModeIsExport
        }
        else
        {
            Test-MSCloudLogin -Platform $Platform `
                -CloudCredential $InboundParameters.GlobalAdminAccount `
                -ConnectionUrl $Url `
                -SkipModuleReload $Global:CurrentModeIsExport
        }
        $data.Add("ConnectionType", "Credential")
        Add-M365DSCTelemetryEvent -Data $data -Type "Connection"
        return "Credential"
    }
    # Case only the ApplicationID and Credentials parameters are specified
    elseif ($null -ne $InboundParameters.GlobalAdminAccount -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId))
    {
        Write-Verbose -Message "GlobalAdminAccount and ApplicationId were specified. Connecting via Delegated Service Principal"
        if ([System.String]::IsNullOrEmpty($url))
        {
            Test-MSCloudLogin -Platform $Platform `
                -ApplicationId $InboundParameters.ApplicationId `
                -CloudCredential $InboundParameters.GlobalAdminAccount `
                -SkipModuleReload $Global:CurrentModeIsExport
        }
        else
        {
            Test-MSCloudLogin -Platform $Platform `
                -ApplicationId $InboundParameters.ApplicationId `
                -CloudCredential $InboundParameters.GlobalAdminAccount `
                -ConnectionUrl $Url `
                -SkipModuleReload $Global:CurrentModeIsExport
        }
        $data.Add("ConnectionType", "ServicePrincipal")
        Add-M365DSCTelemetryEvent -Data $data -Type "Connection"
        return 'ServicePrincipal'
    }
    # Case only the ServicePrincipal with Thumbprint parameters are specified
    elseif ($null -eq $InboundParameters.GlobalAdminAccount -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId) -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.CertificateThumbprint))
    {
        if ([System.String]::IsNullOrEmpty($url))
        {
            Write-Verbose -Message "ApplicationId, TenantId and CertificateThumprint were specified. Connecting via Service Principal"
            Test-MSCloudLogin -Platform $Platform `
                -ApplicationId $InboundParameters.ApplicationId `
                -TenantId $InboundParameters.TenantId `
                -CertificateThumbprint $InboundParameters.CertificateThumbprint `
                -SkipModuleReload $Global:CurrentModeIsExport
        }
        else
        {
            Test-MSCloudLogin -Platform $Platform `
                -ApplicationId $InboundParameters.ApplicationId `
                -TenantId $InboundParameters.TenantId `
                -CertificateThumbprint $InboundParameters.CertificateThumbprint `
                -ConnectionUrl $Url `
                -SkipModuleReload $Global:CurrentModeIsExport
        }
        $data.Add("ConnectionType", "ServicePrincipal")
        Add-M365DSCTelemetryEvent -Data $data -Type "Connection"
        return 'ServicePrincipal'
    }
    # Case only the ServicePrincipal with Thumbprint parameters are specified
    elseif ($null -eq $InboundParameters.GlobalAdminAccount -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId) -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.CertificatePath) -and `
            $null -ne $InboundParameters.CertificatePassword)
    {
        if ([System.String]::IsNullOrEmpty($url))
        {
            Write-Verbose -Message "ApplicationId, TenantId, CertificatePath & CertificatePassword were specified. Connecting via Service Principal"
            Test-MSCloudLogin -Platform $Platform `
                -ApplicationId $InboundParameters.ApplicationId `
                -TenantId $InboundParameters.TenantId `
                -CertificatePassword $InboundParameters.CertificatePassword.Password `
                -CertificatePath $InboundParameters.CertificatePath `
                -SkipModuleReload $Global:CurrentModeIsExport
        }
        else
        {
            Test-MSCloudLogin -Platform $Platform `
                -ApplicationId $InboundParameters.ApplicationId `
                -TenantId $InboundParameters.TenantId `
                -CertificatePassword $InboundParameters.CertificatePassword `
                -CertificatePath $InboundParameters.CertificatePath `
                -ConnectionUrl $Url `
                -SkipModuleReload $Global:CurrentModeIsExport
        }
        $data.Add("ConnectionType", "ServicePrincipal")
        Add-M365DSCTelemetryEvent -Data $data -Type "Connection"
        return 'ServicePrincipal'
    }
    else
    {
        $data.Add("Event", "Error")
        $errorText = 'Unexpected error getting the Authentication Method'
        $data.Add("CustomMessage", $errorText)
        Add-M365DSCTelemetryEvent -Data $data -Type "Error"
        throw $errorText
    }
}

function Get-SPOAdministrationUrl
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [switch]
        $UseMFA,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )
    if ($UseMFA)
    {
        $UseMFASwitch = @{UseMFA = $true }
    }
    else
    {
        $UseMFASwitch = @{ }
    }

    return Get-SPOAdminUrl -CloudCredential $GlobalAdminAccount
}

function Get-M365TenantName
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [switch]
        $UseMFA,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )
    if ($UseMFA)
    {
        $UseMFASwitch = @{UseMFA = $true }
    }
    else
    {
        $UseMFASwitch = @{ }
    }
    Write-Verbose -Message "Connection to Azure AD is required to automatically determine SharePoint Online admin URL..."
    $ConnectionMode = New-M365DSCConnection -Platform 'AzureAD' `
        -InboundParameters $PSBoundParameters
    Write-Verbose -Message "Getting SharePoint Online admin URL..."
    $defaultDomain = Get-AzureADDomain | Where-Object { ($_.Name -like "*.onmicrosoft.com" -or $_.Name -like "*.onmicrosoft.de") -and $_.IsInitial -eq $true } # We don't use IsDefault here because the default could be a custom domain

    if ($defaultDomain[0].Name -like '*.onmicrosoft.com*')
    {
        $tenantName = $defaultDomain[0].Name -replace ".onmicrosoft.com", ""
    }
    elseif ($defaultDomain[0].Name -like '*.onmicrosoft.de*')
    {
        $tenantName = $defaultDomain[0].Name -replace ".onmicrosoft.de", ""
    }

    Write-Verbose -Message "M365 tenant name is $tenantName"
    return $tenantName
}

function Split-ArrayByBatchSize
{
    [OutputType([System.Object[]])]
    Param(
        [Parameter(Mandatory = $true)]
        [System.Object[]]
        $Array,

        [Parameter(Mandatory = $true)]
        [System.Uint32]
        $BatchSize
    )
    for ($i = 0; $i -lt $Array.Count; $i += $BatchSize)
    {
        $NewArray += , @($Array[$i..($i + ($BatchSize - 1))]);
    }
    return $NewArray
}

function Split-ArrayByParts
{
    [OutputType([System.Object[]])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Object[]]
        $Array,

        [Parameter(Mandatory = $true)]
        [System.Uint32]
        $Parts
    )

    if ($Parts)
    {
        $PartSize = [Math]::Ceiling($Array.Count / $Parts)
    }
    $outArray = New-Object 'System.Collections.Generic.List[PSObject]'

    for ($i = 1; $i -le $Parts; $i++)
    {
        $start = (($i - 1) * $PartSize)

        if ($start -lt $Array.Count)
        {
            $end = (($i) * $PartSize) - 1
            if ($end -ge $Array.count)
            {
                $end = $Array.count - 1
            }
            $outArray.Add(@($Array[$start..$end]))
        }
    }
    return , $outArray
}

function Start-DSCInitializedJob
{
    [CmdletBinding()]
    param
    (
        [Parameter()]
        $Name,
        [Parameter(Mandatory = $true)]
        [ScriptBlock]
        $ScriptBlock,
        [Parameter()]
        [Object[]]
        $ArgumentList
    )

    $msloginAssistentPath = (Get-Module MSCloudLoginAssistant).Path.Replace("psm1", "psd1")
    $setupJobScript = " Import-Module '$msloginAssistentPath' -Force | Out-Null;"

    # an explicit import for the teams module because of some problems with missing .net types when used inside a PSJob
    # this only occurs because they are accessed directly, not through a cmdlet from the teams module
    $teamsModuleVersion = (Get-Module MicrosoftTeams).Version
    $setupJobScript += " Import-Module MicrosoftTeams -RequiredVersion $teamsModuleVersion -Force | Out-Null;"

    if ($Global:appIdentityParams)
    {
        $entropyStr = [string]::Join(', ', $Global:appIdentityParams.TokenCacheEntropy)
        $setupJobScript += "[byte[]] `$tokenCacheEntropy = $entropyStr;"
        $setupJobScript += "Init-ApplicationIdentity -Tenant $($Global:appIdentityParams.Tenant) -AppId $($Global:appIdentityParams.AppId) -AppSecret '$($Global:appIdentityParams.AppSecret)' -CertificateThumbprint '$($Global:appIdentityParams.CertificateThumbprint)' -OnBehalfOfUserPrincipalName '$($Global:appIdentityParams.OnBehalfOfUserPrincipalName)' -TokenCacheLocation '$($Global:appIdentityParams.TokenCacheLocation)' -TokenCacheEntropy `$tokenCacheEntropy -TokenCacheDataProtectionScope $($Global:appIdentityParams.TokenCacheDataProtectionScope);"
    }

    # ReverseDSC is needed because most of the time the job will call something from it
    $reverseDscModulePath = (Get-Module ReverseDSC).Path.Replace("psm1", "psd1")
    $setupJobScript += " Import-Module '$reverseDscModulePath' -Force | Out-Null;"


    $insertPosition = 0
    if ($ScriptBlock.Ast.BeginBlock)
    {
        $insertPosition = $ScriptBlock.Ast.BeginBlock.Statements[0].Extent.StartOffset;
    }
    elseif ($ScriptBlock.Ast.ProcessBlock)
    {
        $insertPosition = $ScriptBlock.Ast.ProcessBlock.Statements[0].Extent.StartOffset;
    }
    elseif ($ScriptBlock.Ast.EndBlock)
    {
        $insertPosition = $ScriptBlock.Ast.EndBlock.Statements[0].Extent.StartOffset;
    }
    $insertPosition = $insertPosition - $ScriptBlock.StartPosition.Start
    $strScriptContent = $ScriptBlock.ToString();
    $strScriptContent = $strScriptContent.Insert($insertPosition - 1, $setupJobScript + "`n")
    $newScriptBlock = [ScriptBlock]::Create($strScriptContent)
    Start-Job -Name $Name -ScriptBlock $newScriptBlock  -ArgumentList $ArgumentList
}

function Invoke-M365DSCCommand
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ScriptBlock]
        $ScriptBlock,

        [Parameter()]
        [System.String]
        $InvokationPath,

        [Parameter()]
        [Object[]]
        $Arguments,

        [Parameter()]
        [System.UInt32]
        $Backoff = 2
    )

    $InformationPreference = 'SilentlyContinue'
    $WarningPreference = 'SilentlyContinue'
    $ErrorActionPreference = 'Stop'
    try
    {
        if (-not [System.String]::IsNullOrEmpty($InvokationPath))
        {
            $baseScript = "Import-Module '$InvokationPath\*.psm1' -Force;"
        }

        $invokeArgs = @{
            ScriptBlock = [ScriptBlock]::Create($baseScript + $ScriptBlock.ToString())
        }
        if ($null -ne $Arguments)
        {
            $invokeArgs.Add("ArgumentList", $Arguments)
        }
        return Invoke-Command @invokeArgs
    }
    catch
    {
        if ($_.Exception -like '*M365DSC - *')
        {
            Write-Warning $_.Exception
        }
        else
        {
            if ($Backoff -le 128)
            {
                $NewBackoff = $Backoff * 2
                Write-Warning "    * Throttling detected. Waiting for {$NewBackoff seconds}"
                Start-Sleep -Seconds $NewBackoff
                return Invoke-M365DSCCommand -ScriptBlock $ScriptBlock -Backoff $NewBackoff -Arguments $Arguments -InvokationPath $InvokationPath
            }
            else
            {
                Write-Warning $_
            }
        }
    }
}

function Get-SPOUserProfilePropertyInstance
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $Key,

        [Parameter()]
        [System.String]
        $Value
    )

    $result = [PSCustomObject]@{
        Key   = $Key
        Value = $Value
    }

    return $result
}

function ConvertTo-SPOUserProfilePropertyInstanceString
{
    [CmdletBinding()]
    [OutputType([System.String[]])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Object[]]
        $Properties
    )

    $results = @()
    foreach ($property in $Properties)
    {
        $value = Get-DSCPSTokenValue -Value $property.Value

        $content = "             MSFT_SPOUserProfilePropertyInstance`r`n            {`r`n"
        $content += "                Key   = `"$($property.Key)`"`r`n"
        $content += "                Value = $($value)`r`n"
        $content += "            }`r`n"
        $results += $content
    }
    return $results
}

function Install-M365DSCDevBranch
{
    [CmdletBinding()]
    param()
    #region Download and Extract Dev branch's ZIP
    Write-Host "Downloading the Zip package..." -NoNewline
    $url = "https://github.com/microsoft/Microsoft365DSC/archive/Dev.zip"
    $output = "$($env:Temp)\dev.zip"
    $extractPath = $env:Temp + "\O365Dev"
    Write-Host "Done" -ForegroundColor Green

    Invoke-WebRequest -Uri $url -OutFile $output

    Expand-Archive $output -DestinationPath $extractPath -Force
    #endregion

    #region Install All Dependencies
    $manifest = Import-PowerShellDataFile "$extractPath\Microsoft365DSC-Dev\Modules\Microsoft365DSC\Microsoft365DSC.psd1"
    $dependencies = $manifest.RequiredModules
    foreach ($dependency in $dependencies)
    {
        Write-Host "Installing {$($dependency.ModuleName)}..." -NoNewline
        $existingModule = Get-Module $dependency.ModuleName -ListAvailable | Where-Object -FilterScript { $_.Version -eq $dependency.RequiredVersion }
        if ($null -eq $existingModule)
        {
            Install-Module $dependency.ModuleName -RequiredVersion $dependency.RequiredVersion -Force -AllowClobber | Out-Null
        }
        Import-Module $dependency.ModuleName -Force | Out-Null
        Write-Host "Done" -ForegroundColor Green
    }
    #endregion

    #region Install M365DSC
    Write-Host "Updating the Core Microsoft365DSC module..." -NoNewline
    $defaultPath = 'C:\Program Files\WindowsPowerShell\Modules\Microsoft365DSC\'
    $currentVersionPath = $defaultPath + ([Version]$($manifest.ModuleVersion)).ToString()

    Copy-Item "$extractPath\Microsoft365DSC-Dev\Modules\Microsoft365DSC\*" `
        -Destination $defaultPath -Recurse -Force

    Import-Module ($defaultPath + "Microsoft365DSC.psd1") -Force | Out-Null
    $oldModule = Get-Module 'Microsoft365DSC' | Where-Object -FilterScript { $_.ModuleBase -eq $currentVersionPath }
    Remove-Module $oldModule -Force | Out-Null
    if (Test-Path $currentVersionPath)
    {
        try
        {
            Remove-Item $currentVersionPath -Recurse -Confirm:$false -Force `
                -ErrorAction Stop
        }
        catch
        {
            Write-Verbose $_
        }
    }
    Write-Host "Done" -ForegroundColor Green
    #endregion
}

function Get-AllSPOPackages
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable[]])]
    param(
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword,

        [Parameter()]
        [System.String]
        $CertificateThumbprint
    )

    $ConnectionMode = New-M365DSCConnection -Platform 'PnP' `
        -InboundParameters $PSBoundParameters

    $tenantAppCatalogUrl = Get-PnPTenantAppCatalogUrl

    $ConnectionMode = New-M365DSCConnection -Platform 'PnP' `
        -InboundParameters $PSBoundParameters `
        -Url $tenantAppCatalogUrl

    $filesToDownload = @()

    if ($null -ne $tenantAppCatalogUrl)
    {
        $spfxFiles = Find-PnPFile -List "AppCatalog" -Match '*.sppkg'
        $appFiles = Find-PnPFile -List "AppCatalog" -Match '*.app'

        $allFiles = $spfxFiles + $appFiles

        foreach ($file in $allFiles)
        {
            $filesToDownload += @{Name = $file.Name; Site = $tenantAppCatalogUrl; Title = $file.Title }
        }
    }
    return $filesToDownload
}

function Remove-NullEntriesFromHashtable
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [System.COllections.HashTable]
        $Hash
    )

    $keysToRemove = @()
    foreach ($key in $Hash.Keys)
    {
        if ([System.String]::IsNullOrEmpty($Hash.$key))
        {
            $keysToRemove += $key
        }
    }

    foreach ($key in $keysToRemove)
    {
        $Hash.Remove($key) | Out-Null
    }

    return $Hash
}

# To be deprecated in future release
function Assert-M365DSCTemplate
{
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.String]
        $TemplatePath,

        [Parameter()]
        [System.String]
        $TemplateName
    )
    $InformationPreference = 'SilentlyContinue'
    $WarningPreference = 'SilentlyContinue'

    #region Telemetry
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Event", "AssertTemplate")
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    Write-Host $Global:M365DSCEmojiYellowCircle -NoNewline
    Write-Host " Assert-M365DSCTemplate is deprecated. Please use the new improved Assert-M365DSCBlueprint cmdlet instead." -ForegroundColor Yellow
}

function Assert-M365DSCBlueprint
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $BluePrintUrl,

        [Parameter(Mandatory = $true)]
        [System.String]
        $OutputReportPath,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $Credentials,

        [Parameter()]
        [System.String]
        $HeaderFilePath
    )
    $InformationPreference = 'SilentlyContinue'
    $WarningPreference = 'SilentlyContinue'

    #region Telemetry
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Event", "AssertBlueprint")
    $data.Add("BluePrint", $BluePrintUrl)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $TempBluePrintName = (New-Guid).ToString() + ".M365"
    $LocalBluePrintPath = Join-Path -Path $env:Temp -ChildPath $TempBluePrintName
    try
    {
        # Download the BluePrint locally in a temp location
        Invoke-WebRequest -Uri $BluePrintUrl -OutFile $LocalBluePrintPath
    }
    catch
    {
        # If the download failed, we assume the provided Url was a local path
        # and we try copying the item instead.
        try
        {
            Copy-Item -Path $BluePrintUrl -Destination $LocalBluePrintPath
        }
        catch
        {
            throw $_
        }
    }

    if ((Test-Path -Path $LocalBluePrintPath))
    {
        # Parse the content of the BluePrint into an array of PowerShell Objects
        $parsedBluePrint = ConvertTo-DSCObject -Path $LocalBluePrintPath

        # Generate an Array of Resource Types contained in the BluePrint
        $ResourcesInBluePrint = @()
        foreach ($resource in $parsedBluePrint)
        {
            if ($ResourcesInBluePrint -notcontains $resource.ResourceName)
            {
                $ResourcesInBluePrint += $resource.ResourceName
            }
        }
        Write-Host "Selected BluePrint contains ($($ResourcesInBluePrint.Length)) components to assess."

        # Call the Export-M365DSCConfiguration cmdlet to extract only the resource
        # types contained within the BluePrint;
        Write-Host "Initiating the Export of those ($($ResourcesInBluePrint.Length)) components from the tenant..."
        $TempExportName = (New-Guid).ToString() + ".ps1"
        Export-M365DSCConfiguration -Quiet `
            -ComponentsToExtract $ResourcesInBluePrint `
            -Path $env:temp `
            -FileName $TempExportName `
            -GlobalAdminAccount $Credentials

        # Call the New-M365DSCDeltaReport configuration to generate the Delta Report between
        # the BluePrint and the extracted resources;
        $ExportPath = Join-Path -Path $env:Temp -ChildPath $TempExportName
        New-M365DSCDeltaReport -Source $ExportPath `
            -Destination $LocalBluePrintPath `
            -OutputPath $OutputReportPath `
            -DriftOnly:$true `
            -IsBlueprintAssessment:$true `
            -HeaderFilePath $HeaderFilePath
    }
    else
    {
        Write-Error "M365DSC Template Path {$LocalBluePrintPath} does not exist."
    }
}

function Test-M365DSCDependenciesForNewVersions
{
    [CmdletBinding()]
    $InformationPreference = 'Continue'
    $currentPath = Join-Path -Path $PSScriptRoot -ChildPath '..\' -Resolve
    $manifest = Import-PowerShellDataFile "$currentPath/Microsoft365DSC.psd1"
    $dependencies = $manifest.RequiredModules
    $i = 1
    foreach ($dependency in $dependencies)
    {
        Write-Progress -Activity "Scanning Dependencies" -PercentComplete ($i / $dependencies.Count * 100)
        try
        {
            $moduleInGallery = Find-Module $dependency.ModuleName
            [array]$moduleInstalled = Get-Module $dependency.ModuleName -ListAvailable | Select-Object Version
            $modules = $moduleInstalled | Sort-Object Version -Descending
            $moduleInstalled = $modules[0]
            if ([Version]($moduleInGallery.Version) -gt [Version]($moduleInstalled[0].Version))
            {
                Write-Host "New version of {$($dependency.ModuleName)} is available {$($moduleInGallery.Version)}"
            }
        }
        catch
        {
            Write-Host "New version of {$($dependency.ModuleName)} is available"
        }
        $i++
    }
}

function Update-M365DSCDependencies
{
    [CmdletBinding()]
    $InformationPreference = 'Continue'
    $currentPath = Join-Path -Path $PSScriptRoot -ChildPath '..\' -Resolve
    $manifest = Import-PowerShellDataFile "$currentPath/Microsoft365DSC.psd1"
    $dependencies = $manifest.RequiredModules
    $i = 1
    foreach ($dependency in $dependencies)
    {
        Write-Progress -Activity "Scanning Dependencies" -PercentComplete ($i / $dependencies.Count * 100)
        try
        {
            Install-Module $dependency.ModuleName -RequiredVersion $dependency.RequiredVersion -AllowClobber -Force
        }
        catch
        {
            Write-Host "Could not update {$($dependency.ModuleName)}"
        }
        $i++
    }
}

function Set-M365DSCAgentCertificateConfiguration
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param()

    $existingCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | `
        Where-Object { $_.Subject -match "M365DSCEncryptionCert" }
    if ($null -eq $existingCertificate)
    {
        Write-Verbose -Message "No existing M365DSC certificate found. Creating one."
        $certificateFilePath = "$env:Temp\M365DSC.cer"
        $cert = New-SelfSignedCertificate -Type DocumentEncryptionCertLegacyCsp `
            -DnsName 'Microsoft365DSC' `
            -Subject 'M365DSCEncryptionCert' `
            -HashAlgorithm SHA256 `
            -NotAfter (Get-Date).AddYears(10)
        $cert | Export-Certificate -FilePath $certificateFilePath -Force | Out-Null
        Import-Certificate -FilePath $certificateFilePath `
            -CertStoreLocation 'Cert:\LocalMachine\My' -Confirm:$false | Out-Null
        $existingCertificate = Get-ChildItem -Path Cert:\LocalMachine\My | `
            Where-Object { $_.Subject -match "M365DSCEncryptionCert" }
    }
    else
    {
        Write-Verbose -Message "An existing M365DSc certificate was found. Re-using it."
    }
    $thumbprint = $existingCertificate.Thumbprint
    Write-Verbose -Message "Using M365DSCEncryptionCert with thumbprint {$thumbprint}"

    $configOutputFile = $env:Temp + "\M365DSCAgentLCMConfig.ps1"
    $LCMConfigContent = @"
    [DSCLocalConfigurationManager()]
    Configuration M365AgentConfig
    {
        Node Localhost
        {
            Settings
            {
                CertificateID = '$thumbprint'
            }
        }
    }
    M365AgentConfig | Out-Null
    Set-DSCLocalConfigurationManager M365AgentConfig
"@
    $LCMConfigContent | Out-File $configOutputFile
    & $configOutputFile
    Remove-Item -Path $configOutputFile -Confirm:$false
    Remove-Item -Path "./M365AgentConfig" -Recurse -Confirm:$false
    return $thumbprint
}

function Format-M365ServicePrincipalData
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter()]
        [System.String]
        $configContent,

        [Parameter()]
        [System.String]
        $principal,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint
    )
    if ($configContent.ToLower().Contains($principal.ToLower()))
    {
        $configContent = $configContent -ireplace [regex]::Escape($principal), "`$(`$OrganizationName.Split('.')[0])"
    }
    if ($configContent.ToLower().Contains($ApplicationId.ToLower()))
    {
        $configContent = $configContent -ireplace [regex]::Escape($ApplicationId), "`$(`$ApplicationId)"
    }
    if (-not [System.String]::IsNullOrEmpty($CertificateThumbprint) -and $configContent.ToLower().Contains($CertificateThumbprint.ToLower()))
    {
        $configContent = $configContent -ireplace [regex]::Escape($CertificateThumbprint), "`$(`$CertificateThumbprint)"
    }
    return $configContent
}
function Remove-EmptyValue
{
    [alias('Remove-EmptyValues')]
    [CmdletBinding()]
    param(
        [alias('Splat', 'IDictionary')][Parameter(Mandatory)][System.Collections.IDictionary] $Hashtable,
        [string[]] $ExcludeParameter,
        [switch] $Recursive,
        [int] $Rerun
    )
    foreach ($Key in [string[]] $Hashtable.Keys)
    {
        if ($Key -notin $ExcludeParameter)
        {
            if ($Recursive)
            {
                if ($Hashtable[$Key] -is [System.Collections.IDictionary])
                {
                    if ($Hashtable[$Key].Count -eq 0)
                    {
                        $Hashtable.Remove($Key)
                    }
                    else
                    {
                        Remove-EmptyValue -Hashtable $Hashtable[$Key] -Recursive:$Recursive
                    }
                }
                else
                {
                    if ($null -eq $Hashtable[$Key] -or ($Hashtable[$Key] -is [string] -and $Hashtable[$Key] -eq '') -or ($Hashtable[$Key] -is [System.Collections.IList] -and $Hashtable[$Key].Count -eq 0))
                    {
                        $Hashtable.Remove($Key)
                    }
                }
            }
            else
            {
                if ($null -eq $Hashtable[$Key] -or ($Hashtable[$Key] -is [string] -and $Hashtable[$Key] -eq '') -or ($Hashtable[$Key] -is [System.Collections.IList] -and $Hashtable[$Key].Count -eq 0))
                {
                    $Hashtable.Remove($Key)
                }
            }
        }
    }
    if ($Rerun)
    {
        for ($i = 0; $i -lt $Rerun; $i++)
        {
            Remove-EmptyValue -Hashtable $Hashtable -Recursive:$Recursive
        }
    }
}

function Format-M365ServicePrincipalData
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter()]
        [System.String]
        $configContent,

        [Parameter()]
        [System.String]
        $principal,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint
    )
    if ($configContent.ToLower().Contains($principal.ToLower()))
    {
        $configContent = $configContent -ireplace [regex]::Escape($principal), "`$(`$OrganizationName.Split('.')[0])"
    }
    if ($configContent.ToLower().Contains($ApplicationId.ToLower()))
    {
        $configContent = $configContent -ireplace [regex]::Escape($ApplicationId), "`$(`$ApplicationId)"
    }
    if (-not [System.String]::IsNullOrEmpty($CertificateThumbprint) -and $configContent.ToLower().Contains($CertificateThumbprint.ToLower()))
    {
        $configContent = $configContent -ireplace [regex]::Escape($CertificateThumbprint), "`$(`$CertificateThumbprint)"
    }
    return $configContent
}

function Update-M365DSCExportAuthenticationResults
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        [ValidateSet("Credential", "ServicePrincipal")]
        $ConnectionMode,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $Results
    )

    # we at syskit do not care about the authentication data directly inside the resource
    # because Microsoft365DSC Get-DSCBlock does not handle escaping '$' correctly the default logic works
    # but since we in Trace do handle it corectly within Get-DSCBlockEx the default logic fails and simply cannot work
    # at least not without additional work in the ReverseDSC module
    # removing these properties also results in smaller filesizes for the snapshots
    if ($Results.ContainsKey("ApplicationId"))
    {
        $Results.Remove("ApplicationId") | Out-Null
    }
    if ($Results.ContainsKey("TenantId"))
    {
        $Results.Remove("TenantId") | Out-Null
    }
    if ($Results.ContainsKey("CertificateThumbprint"))
    {
        $Results.Remove("CertificateThumbprint") | Out-Null
    }
    if ($Results.ContainsKey("CertificatePath"))
    {
        $Results.Remove("CertificatePath") | Out-Null
    }
    if ($Results.ContainsKey("CertificatePassword"))
    {
        $Results.Remove("CertificatePassword") | Out-Null
    }
    if ($Results.ContainsKey("GlobalAdminAccount"))
    {
        $Results.Remove("GlobalAdminAccount") | Out-Null
    }

    return $Results


    # default Microsoft365DSC logic

    if ($ConnectionMode -eq 'Credential')
    {
        $Results.GlobalAdminAccount = Resolve-Credentials -UserName "globaladmin"
        if ($Results.ContainsKey("ApplicationId"))
        {
            $Results.Remove("ApplicationId") | Out-Null
        }
        if ($Results.ContainsKey("TenantId"))
        {
            $Results.Remove("TenantId") | Out-Null
        }
        if ($Results.ContainsKey("CertificateThumbprint"))
        {
            $Results.Remove("CertificateThumbprint") | Out-Null
        }
        if ($Results.ContainsKey("CertificatePath"))
        {
            $Results.Remove("CertificatePath") | Out-Null
        }
        if ($Results.ContainsKey("CertificatePassword"))
        {
            $Results.Remove("CertificatePassword") | Out-Null
        }
    }
    else
    {
        if ($Results.ContainsKey("GlobalAdminAccount"))
        {
            $Results.Remove("GlobalAdminAccount") | Out-Null
        }
        if (-not [System.String]::IsNullOrEmpty($Results.ApplicationId))
        {
            $Results.ApplicationId = "`$ConfigurationData.NonNodeData.ApplicationId"
        }
        else
        {
            try
            {
                $Results.Remove("ApplicationId") | Out-Null
            }
            catch
            {
                Write-Verbose -Message "Error removing ApplicationId from Update-M365DSCExportAuthenticationResults"
            }
        }
        if (-not [System.String]::IsNullOrEmpty($Results.CertificateThumbprint))
        {
            $Results.CertificateThumbprint = "`$ConfigurationData.NonNodeData.CertificateThumbprint"
        }
        else
        {
            try
            {
                $Results.Remove("CertificateThumbprint") | Out-Null
            }
            catch
            {
                Write-Verbose -Message "Error removing CertificateThumbprint from Update-M365DSCExportAuthenticationResults"
            }
        }
        if (-not [System.String]::IsNullOrEmpty($Results.CertificatePath))
        {
            $Results.CertificatePath = "`$ConfigurationData.NonNodeData.CertificatePath"
        }
        else
        {
            try
            {
                $Results.Remove("CertificatePath") | Out-Null
            }
            catch
            {
                Write-Verbose -Message "Error removing CertificatePath from Update-M365DSCExportAuthenticationResults"
            }
        }
        if (-not [System.String]::IsNullOrEmpty($Results.TenantId))
        {
            $Results.TenantId = "`$ConfigurationData.NonNodeData.TenantId"
        }
        else
        {
            try
            {
                $Results.Remove("TenantId") | Out-Null
            }
            catch
            {
                Write-Verbose -Message "Error removing TenantId from Update-M365DSCExportAuthenticationResults"
            }
        }
        if ($null -ne $Results.CertificatePassword)
        {
            $Results.CertificatePassword = Resolve-Credentials -UserName "CertificatePassword"
        }
        else
        {
            try
            {
                $Results.Remove("CertificatePassword") | Out-Null
            }
            catch
            {
                Write-Verbose -Message "Error removing CertificatePassword from Update-M365DSCExportAuthenticationResults"
            }
        }
    }
    return $Results
}

function Get-M365DSCExportContentForResource
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $ResourceName,

        [Parameter(Mandatory = $true)]
        [System.String]
        [ValidateSet("Credential", "ServicePrincipal")]
        $ConnectionMode,

        [Parameter(Mandatory = $true)]
        [System.String]
        $ModulePath,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $Results,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount,

        [Parameter()]
        [System.String[]]
        $PropertiesWithDscBlock,

        [Parameter()]
        [System.String[]]
        $PropertiesWithAllowedSpecialCharacters,

        [Parameter()]
        [System.String[]]
        $PropertiesCimArrays,

        [Parameter()]
        [EmbbededResourceInfo[]]
        $PropertiesWithEmbeddedResources
    )

    if ($Results -and $Results.Ensure -and $Results.Ensure -eq "Absent")
    {
        return ""
    }

    $OrganizationName = ""
    if ($ConnectionMode -eq 'ServicePrincipal')
    {
        $OrganizationName = $TenantId
    }
    else
    {
        $OrganizationName = $GlobalAdminAccount.UserName.Split('@')[1]
    }

    # Ensure the string properties are properly formatted;

    # This is the M365DSC fix for special characters. We at SysKit have are own inside our own Get-DSCBlockEx and for the timebeing will stick with it
    # $Results = Format-M365DSCString -Properties $Results `
    #     -ResourceName $ResourceName

    $content = "        $ResourceName " + (New-Guid).ToString() + "`r`n"
    $content += "        {`r`n"
    $partialContent = Get-DSCBlockEx -Params $Results -ModulePath $ModulePath -PropertiesWithAllowedSpecialCharacters $PropertiesWithAllowedSpecialCharacters -PropertiesWithEmbeddedResources $PropertiesWithEmbeddedResources
    if ($ConnectionMode -eq 'Credential')
    {
        $partialContent = Convert-DSCStringParamToVariable -DSCBlock $partialContent `
            -ParameterName "GlobalAdminAccount"
    }
    else
    {
        if (![System.String]::IsNullOrEmpty($Results.ApplicationId))
        {
            $partialContent = Convert-DSCStringParamToVariable -DSCBlock $partialContent `
                -ParameterName "ApplicationId"
        }
        if (![System.String]::IsNullOrEmpty($Results.TenantId))
        {
            $partialContent = Convert-DSCStringParamToVariable -DSCBlock $partialContent `
                -ParameterName "TenantId"
        }
        if (![System.String]::IsNullOrEmpty($Results.CertificatePath))
        {
            $partialContent = Convert-DSCStringParamToVariable -DSCBlock $partialContent `
                -ParameterName "CertificatePath"
        }
        if (![System.String]::IsNullOrEmpty($Results.CertificateThumbprint))
        {
            $partialContent = Convert-DSCStringParamToVariable -DSCBlock $partialContent `
                -ParameterName "CertificateThumbprint"
        }
        if (![System.String]::IsNullOrEmpty($Results.CertificatePassword))
        {
            $partialContent = Convert-DSCStringParamToVariable -DSCBlock $partialContent `
                -ParameterName "CertificatePassword"
        }
    }

    foreach ($dscProp in $PropertiesWithDscBlock)
    {
        $isCimArray = $null -ne $PropertiesCimArrays -and $PropertiesCimArrays.Contains($dscProp)
        $partialContent = Convert-DSCStringParamToVariable -DSCBlock $partialContent `
            -ParameterName $dscProp -IsCIMArray $isCimArray
    }

    if ($partialContent.ToLower().IndexOf($OrganizationName.ToLower()) -gt 0)
    {
        $partialContent = $partialContent -ireplace [regex]::Escape($OrganizationName + ":"), "`$($OrganizationName):"
        $partialContent = $partialContent -ireplace [regex]::Escape($OrganizationName), "`$OrganizationName"
        $partialContent = $partialContent -ireplace [regex]::Escape("@" + $OrganizationName), "@`$OrganizationName"
    }
    $content += $partialContent
    $content += "        }`r`n"
    return $content
}

function Test-M365DSCNewVersionAvailable
{
    [CmdletBinding()]
    param()

    try
    {
        if ($null -eq $Global:M365DSCNewVersionNotification)
        {
            # Get current module used
            $currentVersion = Get-Module 'Microsoft365DSC' -ErrorAction Stop

            # Get module in the Gallery
            $JobID = Start-Job { Find-Module 'Microsoft365DSC' -ErrorAction Stop }
            $Timeout = $true
            for ($i = 0; $i -lt 10; $i++)
            {
                if ((Get-Job $JobID.id).State -notmatch 'Running')
                {
                    $Timeout = $false
                    break;
                }
                Start-Sleep -Seconds 1
            }
            if ($Timeout)
            {
                return
            }
            $GalleryVersion = Get-Job $JobID.id | Receive-Job
            if ([Version]($GalleryVersion.Version) -gt [Version]($currentVersion.Version))
            {
                $message = "A NEWER VERSION OF MICROSOFT365DSC {v$($GalleryVersion.Version)} IS AVAILABLE IN THE POWERSHELL GALLERY. TO UPDATE, RUN:`r`nInstall-Module Microsoft365DSC -Force -AllowClobber"
                Write-Host $message `
                    -ForegroundColor 'White' `
                    -BackgroundColor 'DarkGray'
                Write-Verbose -Message $message
            }
            $Global:M365DSCNewVersionNotification = 'AlreadyShown'
        }
    }
    catch
    {
        Write-Verbose -Message $_
        Add-M365DSCEvent -Message $_ -EntryType 'Error' `
            -EventID 1 -Source $($MyInvocation.MyCommand.Source)
    }
}

function Execute-CSOMQueryRetry
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.ClientContext] $context
    )

    [Microsoft.SharePoint.Client.ClientContextExtensions]::ExecuteQueryRetry($context)
}

<#
.Synopsis
    Facilitates the loading of specific properties of a Microsoft.SharePoint.Client.ClientObject object or Microsoft.SharePoint.Client.ClientObjectCollection object.
.DESCRIPTION
    Replicates what you would do with a lambda expression in C#.
    For example, "ctx.Load(list, l => list.Title, l => list.Id)" becomes
    "Load-CSOMProperties -object $list -propertyNames @('Title', 'Id')".
.EXAMPLE
    Load-CSOMProperties -parentObject $web -collectionObject $web.Fields -propertyNames @("InternalName", "Id") -parentPropertyName "Fields" -executeQuery
    $web.Fields | select InternalName, Id
.EXAMPLE
   Load-CSOMProperties -object $web -propertyNames @("Title", "Url", "AllProperties") -executeQuery
   $web | select Title, Url, AllProperties
#>
function Load-CSOMProperties
{
    [CmdletBinding(DefaultParameterSetName = 'ClientObject')]
    param (
        # The Microsoft.SharePoint.Client.ClientObject to populate.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0, ParameterSetName = "ClientObject")]
        [Microsoft.SharePoint.Client.ClientObject]
        $object,

        # The Microsoft.SharePoint.Client.ClientObject that contains the collection object.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0, ParameterSetName = "ClientObjectCollection")]
        [Microsoft.SharePoint.Client.ClientObject]
        $parentObject,

        # The Microsoft.SharePoint.Client.ClientObjectCollection to populate.
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1, ParameterSetName = "ClientObjectCollection")]
        [Microsoft.SharePoint.Client.ClientObjectCollection]
        $collectionObject,

        # The object properties to populate
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "ClientObject")]
        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "ClientObjectCollection")]
        [string[]]
        $propertyNames,

        # The parent object's property name corresponding to the collection object to retrieve (this is required to build the correct lamda expression).
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = "ClientObjectCollection")]
        [string]
        $parentPropertyName,

        # If specified, execute the ClientContext.ExecuteQuery() method.
        [Parameter(Mandatory = $false, Position = 4)]
        [switch]
        $executeQuery
    )

    begin
    {
    }
    process
    {
        if ($PsCmdlet.ParameterSetName -eq "ClientObject")
        {
            $type = $object.GetType()
        }
        else
        {
            $type = $collectionObject.GetType()
            if ($collectionObject -is [Microsoft.SharePoint.Client.ClientObjectCollection])
            {
                $type = $collectionObject.GetType().BaseType.GenericTypeArguments[0]
            }
        }

        $exprType = [System.Linq.Expressions.Expression]
        $parameterExprType = [System.Linq.Expressions.ParameterExpression].MakeArrayType()
        $lambdaMethod = $exprType.GetMethods() | Where-Object { $_.Name -eq "Lambda" -and $_.IsGenericMethod -and $_.GetParameters().Length -eq 2 -and $_.GetParameters()[1].ParameterType -eq $parameterExprType }
        $lambdaMethodGeneric = Invoke-Expression "`$lambdaMethod.MakeGenericMethod([System.Func``2[$($type.FullName),System.Object]])"
        $expressions = @()

        foreach ($propertyName in $propertyNames)
        {
            $param1 = [System.Linq.Expressions.Expression]::Parameter($type, "p")
            try
            {
                $name1 = [System.Linq.Expressions.Expression]::Property($param1, $propertyName)
            }
            catch
            {
                Write-Error "Instance property '$propertyName' is not defined for type $type"
                return
            }
            $body1 = [System.Linq.Expressions.Expression]::Convert($name1, [System.Object])
            $expression1 = $lambdaMethodGeneric.Invoke($null, [System.Object[]] @($body1, [System.Linq.Expressions.ParameterExpression[]] @($param1)))

            if ($collectionObject -ne $null)
            {
                $expression1 = [System.Linq.Expressions.Expression]::Quote($expression1)
            }
            $expressions += @($expression1)
        }


        if ($PsCmdlet.ParameterSetName -eq "ClientObject")
        {
            $object.Context.Load($object, $expressions)
            if ($executeQuery)
            { $object.Context.ExecuteQuery()
            }
        }
        else
        {
            $newArrayInitParam1 = Invoke-Expression "[System.Linq.Expressions.Expression``1[System.Func````2[$($type.FullName),System.Object]]]"
            $newArrayInit = [System.Linq.Expressions.Expression]::NewArrayInit($newArrayInitParam1, $expressions)

            $collectionParam = [System.Linq.Expressions.Expression]::Parameter($parentObject.GetType(), "cp")
            $collectionProperty = [System.Linq.Expressions.Expression]::Property($collectionParam, $parentPropertyName)

            $expressionArray = @($collectionProperty, $newArrayInit)
            $includeMethod = [Microsoft.SharePoint.Client.ClientObjectQueryableExtension].GetMethod("Include")
            $includeMethodGeneric = Invoke-Expression "`$includeMethod.MakeGenericMethod([$($type.FullName)])"

            $lambdaMethodGeneric2 = Invoke-Expression "`$lambdaMethod.MakeGenericMethod([System.Func``2[$($parentObject.GetType().FullName),System.Object]])"
            $callMethod = [System.Linq.Expressions.Expression]::Call($null, $includeMethodGeneric, $expressionArray)

            $expression2 = $lambdaMethodGeneric2.Invoke($null, @($callMethod, [System.Linq.Expressions.ParameterExpression[]] @($collectionParam)))

            $parentObject.Context.Load($parentObject, $expression2)
            if ($executeQuery)
            { $parentObject.Context.ExecuteQuery()
            }
        }
    }
    end
    {
    }
}

function Get-DSCPSTokenValue
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter()]
        [Object]
        $Value,

        [Parameter()]
        [bool]
        $AllowSpecialCharachters
    )

    if ($null -eq $Value)
    {
        $Value = ""
    }

    if (!($Value -is [string]))
    {
        $Value = $Value.ToString()
    }

    $sb = [System.Text.StringBuilder]::new()

    [void]$sb.Append("`"")

    if (!$AllowSpecialCharachters)
    {
        foreach ($char in [char[]]$Value)
        {
            if ($char -eq '"' -or $char -eq [char]0x201C -or $char -eq [char]0x201D -or $char -eq [char]0x201E -or $char -eq '$')
            {
                [void]$sb.Append('`')
            }
            [void]$sb.Append($char)
        }
    }
    else
    {
        [void]$sb.Append($Value);
    }

    [void]$sb.Append("`"")
    $retVal = $sb.ToString()
    return $retVal
}

class EmbbededResourceInfo
{
    [string]$PropertyName
    [string]$ResourceName
}

function Get-DSCBlockEx
{
    <#
.SYNOPSIS
Generate the DSC string representing the resource's instance.

.DESCRIPTION
This function is really the core of ReverseDSC. It takes in an array of
parameters and returns the DSC string that represents the given instance
of the specified resource.

.PARAMETER ModulePath
Full file path to the .psm1 module we are looking to get an instance of.
In most cases this will be the full path to the .psm1 file of the DSC resource.

.PARAMETER Params
Hashtable that contains the list of Key properties and their values.

#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $ModulePath,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $Params,

        [Parameter()]
        [System.String[]]
        $PropertiesWithAllowedSpecialCharacters,

        [Parameter()]
        [EmbbededResourceInfo[]]
        $PropertiesWithEmbeddedResources,

        [Parameter()]
        [int]
        $IndentValue = 0
    )

    $Sorted = $Params.GetEnumerator() | Sort-Object -Property Name
    $NewParams = [Ordered]@{}

    foreach ($entry in $Sorted)
    {
        $NewParams.Add($entry.Key, $entry.Value)
    }

    # Figure out what parameter has the longuest name, and get its Length;
    $maxParamNameLength = 0
    foreach ($param in $NewParams.Keys)
    {
        if ($param.Length -gt $maxParamNameLength)
        {
            $maxParamNameLength = $param.Length
        }
    }

    # PSDscRunAsCredential is 20 characters and in most case the longuest.
    if ($maxParamNameLength -lt 20)
    {
        $maxParamNameLength = 20
    }

    $dscBlock = ""
    $NewParams.Keys | ForEach-Object {
        if ($null -ne $NewParams[$_])
        {
            $paramType = $NewParams[$_].GetType().Name
        }
        else
        {
            $paramType = Get-DSCParamType -ModulePath $ModulePath -ParamName "`$$_"
        }

        $AllowSpecialCharachtersInValues = $null -ne $PropertiesWithAllowedSpecialCharacters -and $PropertiesWithAllowedSpecialCharacters.Contains($_) -or $_ -eq 'GlobalAdminAccount'

        $value = $null
        $propName = $_
        if ($null -ne $PropertiesWithEmbeddedResources -and $PropertiesWithEmbeddedResources.Length -gt 0 -and ($idx = [System.Array]::FindIndex($PropertiesWithEmbeddedResources, [Predicate[EmbbededResourceInfo]] { $args[0].PropertyName -eq $propName })) -ge 0)
        {
            $meta = $PropertiesWithEmbeddedResources[$idx]
            if ($paramType -eq "ArrayList" -or $paramType -eq "List``1" -or $paramType -eq "Hashtable[]")
            {
                $value = "@("
                foreach ($obj in $NewParams[$_])
                {
                    $value += "                  $($meta.ResourceName)`r`n            {`r`n"
                    $value += Get-DSCBlockEx -ModulePath $ModulePath -Params $obj -IndentValue 4
                    $value += "            }`r`n"
                }
                $value += "            )";
            }
        }
        elseif ($paramType -eq "System.String" -or $paramType -eq "String" -or $paramType -eq "Guid" -or $paramType -eq 'TimeSpan')
        {
            $value = Get-DSCPSTokenValue -Value ($NewParams.Item($_)) -AllowSpecialCharachters $AllowSpecialCharachtersInValues
        }
        elseif ($paramType -eq "System.Boolean" -or $paramType -eq "Boolean")
        {
            $value = "`$" + $NewParams.Item($_)
        }
        elseif ($paramType -eq "System.Management.Automation.PSCredential")
        {
            if ($null -ne $NewParams.Item($_))
            {
                if ($NewParams.Item($_).ToString() -like "`$Creds*")
                {
                    $value = $NewParams.Item($_).Replace("-", "_").Replace(".", "_")
                }
                else
                {
                    if ($null -eq $NewParams.Item($_).UserName)
                    {
                        $value = "`$Creds" + ($NewParams.Item($_).Split('\'))[1].Replace("-", "_").Replace(".", "_")
                    }
                    else
                    {
                        if ($NewParams.Item($_).UserName.Contains("@") -and !$NewParams.Item($_).UserName.COntains("\"))
                        {
                            $value = "`$Creds" + ($NewParams.Item($_).UserName.Split('@'))[0]
                        }
                        else
                        {
                            $value = "`$Creds" + ($NewParams.Item($_).UserName.Split('\'))[1].Replace("-", "_").Replace(".", "_")
                        }
                    }
                }
            }
            else
            {
                $value = "Get-Credential -Message " + $_
            }
        }
        elseif ($paramType -eq "System.Collections.Hashtable" -or $paramType -eq "Hashtable")
        {
            $value = "@{"
            $hash = $NewParams.Item($_)
            $hash.Keys | ForEach-Object {
                try
                {
                    $escapedItemValue = Get-DSCPSTokenValue -Value $hash.Item($_) -AllowSpecialCharachters $AllowSpecialCharachtersInValues
                    $value += $_ + " = " + $escapedItemValue + "; "
                }
                catch
                {
                    $value = $hash
                }
            }
            $value += "}"
        }
        elseif ($paramType -eq "System.String[]" -or $paramType -eq "String[]" -or $paramType -eq "ArrayList" -or $paramType -eq "List``1")
        {
            $hash = $NewParams.Item($_)
            if ($hash -and !$hash.ToString().StartsWith("`$ConfigurationData."))
            {
                $value = "@("
                $hash | ForEach-Object {
                    $escapedValue = Get-DSCPSTokenValue -Value $_ -AllowSpecialCharachters $AllowSpecialCharachtersInValues
                    $value += $escapedValue + ","
                }
                if ($value.Length -gt 2)
                {
                    $value = $value.Substring(0, $value.Length - 1)
                }
                $value += ")"
            }
            else
            {
                if ($hash)
                {
                    $value = $hash
                }
                else
                {
                    $value = "@()"
                }
            }
        }
        elseif ($paramType -eq "System.UInt32[]")
        {
            $hash = $NewParams.Item($_)
            if ($hash)
            {
                $value = "@("
                $hash | ForEach-Object {
                    $value += $_.ToString() + ","
                }
                if ($value.Length -gt 2)
                {
                    $value = $value.Substring(0, $value.Length - 1)
                }
                $value += ")"
            }
            else
            {
                if ($hash)
                {
                    $value = $hash
                }
                else
                {
                    $value = "@()"
                }
            }
        }
        elseif ($paramType -eq "Object[]" -or $paramType -eq "Microsoft.Management.Infrastructure.CimInstance[]")
        {
            $array = $hash = $NewParams.Item($_)

            if ($array.Length -gt 0 -and ($array[0].GetType().Name -eq "String" -and $paramType -ne "Microsoft.Management.Infrastructure.CimInstance[]"))
            {
                $value = "@("
                $hash | ForEach-Object {
                    $escapedItemValue = Get-DSCPSTokenValue -Value $_ -AllowSpecialCharachters $AllowSpecialCharachtersInValues
                    $value += $escapedItemValue + ","
                }
                if ($value.Length -gt 2)
                {
                    $value = $value.Substring(0, $value.Length - 1)
                }
                $value += ")"
            }
            else
            {
                $value = "@("
                $array | ForEach-Object {
                    $value += $_
                }
                $value += ")"
            }
        }
        elseif ($paramType -eq "CimInstance")
        {
            $value = $NewParams[$_]
        }
        else
        {
            if ($null -eq $NewParams[$_])
            {
                $value = "`$null"
            }
            else
            {
                if ($NewParams[$_].GetType().BaseType.Name -eq "Enum")
                {
                    $value = "`"" + $NewParams.Item($_) + "`""
                }
                else
                {
                    $value = $NewParams.Item($_)
                }
            }
        }

        # Determine the number of additional spaces we need to add before the '=' to make sure the values are all aligned. This number
        # is obtained by substracting the length of the current parameter's name to the maximum length found.
        $numberOfAdditionalSpaces = $maxParamNameLength - $_.Length
        $additionalSpaces = ""
        for ($i = 0; $i -lt $numberOfAdditionalSpaces; $i++)
        {
            $additionalSpaces += " "
        }
        $dscBlock += [string]::new(' ', 12 + $IndentValue) + $_ + $additionalSpaces + " = " + $value + ";`r`n"
    }

    return $dscBlock
}
function Get-M365DSCComponentsForAuthenticationType
{
    [CmdletBinding()]
    [OutputType([System.String[]])]
    param(
        [Parameter()]
        [System.String[]]
        [ValidateSet('Application', 'Certificate', 'Credentials')]
        $AuthenticationMethod
    )

    $modules = Get-ChildItem -Path ($PSScriptRoot + "\..\DSCResources\") -Recurse -Filter '*.psm1'
    $Components = @()
    foreach ($resource in $modules)
    {
        Import-Module $resource.FullName -Force
        $parameters = (Get-Command 'Set-TargetResource').Parameters.Keys

        # Case - Resource only supports AppID & GlobalAdmin
        if ($AuthenticationMethod.Contains("Application") -and `
                $AuthenticationMethod.Contains("Credentials") -and `
            ($parameters.Contains("ApplicationId") -and `
                    $parameters.Contains("GlobalAdminAccount") -and `
                    -not $parameters.Contains('CertificateThumbprint') -and `
                    -not $parameters.Contains('CertificatePath') -and `
                    -not $parameters.Contains('CertificatePassword') -and `
                    -not $parameters.Contains('TenantId')))
        {
            $Components += $resource.Name.Replace("MSFT_", "").Replace(".psm1", "")
        }

        #Case - Resource certificate info and TenantId
        elseif ($AuthenticationMethod.Contains("Certificate") -and `
            ($parameters.Contains('CertificateThumbprint') -or `
                    $parameters.Contains('CertificatePath') -or `
                    $parameters.Contains('CertificatePassword')) -and `
                $parameters.Contains('TenantId'))
        {
            $Components += $resource.Name.Replace("MSFT_", "").Replace(".psm1", "")
        }

        # Case - Resource contains GlobalAdminAccount
        elseif ($AuthenticationMethod.Contains("Credentials") -and `
                $parameters.Contains('GlobalAdminAccount'))
        {
            $Components += $resource.Name.Replace("MSFT_", "").Replace(".psm1", "")
        }
    }
    return $Components
}

function Get-M365DSCAllResources
{
    [CmdletBinding()]
    [OutputType([System.String[]])]
    [CmdletBinding()]
    param ()

    $allResources = Get-ChildItem -Path ($PSScriptRoot + "\..\DSCResources\") -Recurse -Filter '*.psm1'
    $result = @()
    foreach ($resource in $allResources)
    {
        $result += $resource.Name.Replace("MSFT_", "").Replace(".psm1", "")
    }

    return $result
}

function Test-M365DSCObjectHasProperty
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true, Position = 1)]
        [Object]
        $Object,

        [Parameter(Mandatory = $true, Position = 2)]
        [String]
        $PropertyName
    )

    if (([bool]($Object.PSobject.Properties.name -contains $PropertyName)) -eq $true)
    {
        if ($null -ne $Object.$PropertyName)
        {
            return $true
        }
    }
    return $false
}
