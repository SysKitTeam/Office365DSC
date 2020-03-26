function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Url,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Identity,

        [Parameter()]
        [System.String]
        $Owner,

        [Parameter()]
        [System.String[]]
        $PermissionLevels,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount,

        [Parameter()]
        $RawInputObject
    )

    Write-Verbose -Message "Getting SPOSiteGroups for {$Url}"
    #region Telemetry
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $MyInvocation.MyCommand.ModuleName)
    $data.Add("Method", $MyInvocation.MyCommand)
    Add-O365DSCTelemetryEvent -Data $data
    #endregion

    $nullReturn = @{
        Url                = $Url
        Identity           = $null
        Owner              = $null
        PermissionLevels   = $null
        GlobalAdminAccount = $GlobalAdminAccount
        Ensure             = "Absent"
    }


    if ($RawInputObject)
    {
        $siteGroup = $RawInputObject.SiteGroup
        $sitePermissions = $RawInputObject.SitePermissions
    }
    else
    {
        Test-MSCloudLogin -CloudCredential $GlobalAdminAccount `
        -Platform PnP
        #checking if the site actually exists
        try
        {
            $site = Get-PnPTenantSite $Url
        }
        catch
        {
            $Message = "The specified site collection doesn't exist."
            New-Office365DSCLogEntry -Error $_ -Message $Message
            throw $Message
        }
        try
        {
            Test-MSCloudLogin -CloudCredential $GlobalAdminAccount `
                -Platform PnP `
                -ConnectionUrl $Url
            $siteGroup = Get-PnPGroup -Identity $Identity
        }
        catch
        {
            if ($Error[0].Exception.Message -eq "Group cannot be found.")
            {
                write-verbose -Message "Site group $($Identity) could not be found on site $($Url)"

            }
        }
        if ($null -eq $siteGroup)
        {
            return $nullReturn
        }

        try
        {
            $sitePermissions = Get-PnPGroupPermissions -Identity $Identity -ErrorAction Stop
        }
        catch
        {
            if ($_.Exception -like '*Access denied*')
            {
                Write-Warning -Message "The specified account does not have access to the permissions list for {$Url}"
                return $nullReturn
            }
        }
    }
    $permissions = @()
    foreach ($entry in $sitePermissions.RoleTypeKind)
    {
        $permissions += $entry.ToString()
    }
    return @{
        Url                = $Url
        Identity           = $siteGroup.Title
        Owner              = $siteGroup.Owner.LoginName
        PermissionLevels   = $permissions
        GlobalAdminAccount = $GlobalAdminAccount
        Ensure             = "Present"
    }
}

function Set-TargetResource
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Url,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Identity,

        [Parameter()]
        [System.String]
        $Owner,

        [Parameter()]
        [System.String[]]
        $PermissionLevels,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )

    Write-Verbose -Message "Setting SPOSiteGroups for {$Url}"
    #region Telemetry
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $MyInvocation.MyCommand.ModuleName)
    $data.Add("Method", $MyInvocation.MyCommand)
    Add-O365DSCTelemetryEvent -Data $data
    #endregion

    Test-MSCloudLogin -CloudCredential $GlobalAdminAccount `
        -Platform PnP `
        -ErrorAction SilentlyContinue

    $currentValues = Get-TargetResource @PSBoundParameters
    if ($Ensure -eq "Present" -and $currentValues.Ensure -eq "Absent")
    {
        $SiteGroupSettings = @{
            Title = $Identity
            Owner = $Owner
        }
        Write-Verbose -Message "Site group $Identity does not exist, creating it."
        New-PnPGroup @SiteGroupSettings
    }
    elseif ($Ensure -eq "Present" -and $currentValues.Ensure -eq "Present")
    {
        $RefferenceObjectRoles = $PermissionLevels
        $DifferenceObjectRoles = $currentValues.PermissionLevels
        $compareOutput = Compare-Object -ReferenceObject $RefferenceObjectRoles -DifferenceObject $DifferenceObjectRoles
        $PermissionLevelsToAdd = @()
        $PermissionLevelsToRemove = @()
        foreach ($entry in $compareOutput)
        {
            if ($entry.SideIndicator -eq "<=")
            {
                Write-Verbose -Message "Permissionlevels to add: $($entry.InputObject)"
                $PermissionLevelsToAdd += $entry.InputObject
            }
            else
            {
                Write-Verbose -Message "Permissionlevels to remove: $($entry.InputObject)"
                $PermissionLevelsToRemove += $entry.InputObject
            }
        }
        if ($PermissionLevelsToAdd.Count -eq 0 -and $PermissionLevelsToRemove.Count -ne 0)
        {
            $SiteGroupSettings = @{
                Identity = $Identity
                Owner    = $Owner
            }
            Set-PnPGroup @SiteGroupSettings

            Set-PnPGroupPermissions -Identity $Identity -RemoveRole $PermissionLevelsToRemove
        }
        elseif ($PermissionLevelsToRemove.Count -eq 0 -and $PermissionLevelsToAdd.Count -ne 0)
        {
            $SiteGroupSettings = @{
                Identity = $Identity
                Owner    = $Owner
            }
            Set-PnPGroup @SiteGroupSettings

            Set-PnPGroupPermissions -Identity $Identity -AddRole $PermissionLevelsToAdd
        }
        elseif ($PermissionLevelsToAdd.Count -eq 0 -and $PermissionLevelsToRemove.Count -eq 0)
        {
            if (($Identity -eq $currentValues.Identity) -and ($Owner -eq $currentlValues.Owner))
            {
                Write-Verbose -Message "All values are configured as desired"
            }
            else
            {
                $SiteGroupSettings = @{
                    Identity = $Identity
                    Owner    = $Owner
                }
                Set-PnPGroup @SiteGroupSettings
            }
        }
        else
        {
            $SiteGroupSettings = @{
                Identity = $Identity
                Owner    = $Owner
            }
            Set-PnPGroup @SiteGroupSettings

            Set-PnPGroupPermissions -Identity $Identity -AddRole $PermissionLevelsToAdd -RemoveRole $PermissionLevelsToRemove
        }

    }
    elseif ($Ensure -eq "Absent" -and $currentValues.Ensure -eq "Present")
    {
        $SiteGroupSettings = @{
            Identity = $Identity
        }
        Write-Verbose "Removing SPOSiteGroup $Identity"
        Remove-PnPGroup @SiteGroupSettings
    }
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Url,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Identity,

        [Parameter()]
        [System.String]
        $Owner,

        [Parameter()]
        [System.String[]]
        $PermissionLevels,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = "Present",

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )

    Write-Verbose -Message "Testing SPOSiteGroups for {$Url}"
    $CurrentValues = Get-TargetResource @PSBoundParameters

    Write-Verbose -Message "Current Values: $(Convert-O365DscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-O365DscHashtableToString -Hashtable $PSBoundParameters)"

    $ValuesToCheck = $PSBoundParameters
    $ValuesToCheck.Remove('GlobalAdminAccount') | Out-Null

    $TestResult = Test-Office365DSCParameterState -CurrentValues $CurrentValues `
        -Source $($MyInvocation.MyCommand.Source) `
        -DesiredValues $PSBoundParameters `
        -ValuesToCheck $ValuesToCheck.Keys

    Write-Verbose -Message "Test-TargetResource returned $TestResult"

    return $TestResult
}

function Export-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )
    #region Telemetry
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $MyInvocation.MyCommand.ModuleName)
    $data.Add("Method", $MyInvocation.MyCommand)
    Add-O365DSCTelemetryEvent -Data $data
    #endregion

    $InformationPreference = 'Continue'
    Test-MSCloudLogin -CloudCredential $GlobalAdminAccount `
        -Platform PnP `
        -ErrorAction SilentlyContinue

    #Loop through all sites
    #for each site loop through all site groups and retrieve parameters
    $sites = Get-PnPTenantSite

    $i = 1
    $content = ""
    foreach ($site in $sites)
    {
        Write-Information "    [$i/$($sites.Length)] SPOSite groups for {$($site.Url)}"
        try
        {
            Test-MSCloudLogin -CloudCredential $GlobalAdminAccount `
                -Platform PnP `
                -ConnectionUrl $site.Url
            $siteGroups = Get-PnPGroup
        }
        catch
        {
            $message = $Error[0].Exception.Message
            if ($null -ne $message)
            {
                Write-Warning -Message $message
            }
            else
            {
                Write-Verbose -Message "Could not retrieve sitegroups for site $($site.Url)"
            }
        }
        foreach ($siteGroup in $siteGroups)
        {
            try
            {
                $sitePerm = Get-PnPGroupPermissions -Identity $siteGroup.Title -ErrorAction Stop
            }
            catch
            {
                Write-Warning -Message "The specified account does not have access to the permissions list for {$Url}"
                break
            }
            $params = @{
                Url                = $site.Url
                Identity           = $siteGroup.Title
                GlobalAdminAccount = $GlobalAdminAccount
                RawInputObject     = @{
                    Site = $site
                    SiteGroup = $siteGroup
                    SitePermissions = $sitePerm
                }
            }
            try
            {
                $result = Get-TargetResource @params
                $result = Remove-NullEntriesFromHashtable -Hash $result
                $result.GlobalAdminAccount = Resolve-Credentials -UserName "globaladmin"
                $content += "        SPOSiteGroup " + (New-GUID).ToString() + "`r`n"
                $content += "        {`r`n"
                $currentDSCBlock = Get-DSCBlockEx -Params $result -ModulePath $PSScriptRoot
                $content += Convert-DSCStringParamToVariable -DSCBlock $currentDSCBlock -ParameterName "GlobalAdminAccount"
                $content += "        }`r`n"
            }
            catch
            {
                Write-Verbose "There was an issue retrieving the SiteGroups for $($Url)"
            }
        }

        $i++
    }
    return $content
}

Export-ModuleMember -Function *-TargetResource
