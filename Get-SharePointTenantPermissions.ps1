# Get-SharePointTenantPermissions.ps1
# Description: This script will get all the permissions for a given user or users in a SharePoint Online tenant and export them to a CSV file.

#requires -Modules PnP.PowerShell, MSAL.PS
param (
    [Parameter(Mandatory = $true)]
    [string] $TenantName,
    [Parameter(Mandatory = $true)]
    [string] $UserEmail,
    [Parameter(Mandatory = $true)]
    [string] $CSVPath,
    [Parameter(Mandatory = $true)]
    [string] $ClientId,
    [Parameter(Mandatory = $true)]
    [string] $CertificatePath,
    [Parameter(Mandatory = $false)]
    [SecureString] $CertificatePassword,
    [Parameter(Mandatory = $false)]
    [int] $ThrottleLimit = 12,
    [Parameter(Mandatory = $false)]
    [switch] $Append = $false,
    [Parameter(Mandatory = $false)]
    [string] $Log,
    [Parameter(Mandatory = $false)]
    [switch] $AppendLog
)

# Start benchmarking for this user
$scriptStartTime = Get-Date

function Connect-TenantSite {
    <#
    .SYNOPSIS
    Connects to a SharePoint Online site using certificate-based authentication via PnP PowerShell.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $SiteUrl
    )

    $connectionAttempts = 3
    for ($i = 0; $i -lt $connectionAttempts; $i++) {
        try {
            if ($CertificatePassword) {
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -Tenant "$TenantName.onmicrosoft.com"
            }
            else {
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -CertificatePath $CertificatePath -Tenant "$TenantName.onmicrosoft.com"
            }
            break
        }
        catch {
            if ($i -eq $connectionAttempts - 1) {
                Write-Error $_.Exception.Message
                throw $_
            }
            continue
        }
    }
}

function Get-GraphToken {
    <#
    .SYNOPSIS
    Gets a bearer token for the Microsoft Graph API using certificate-based authentication.
    #>
    # Build certificate object from PFX, using password if provided
    if ($CertificatePassword) {
        $plainPassword = [System.Net.NetworkCredential]::new('', $CertificatePassword).Password
        $certObject = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $plainPassword)
    }
    else {
        $certObject = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
    }

    $connectionParameters = @{
        'TenantId'          = "$TenantName.onmicrosoft.com"
        'ClientId'          = $ClientId
        'ClientCertificate' = $certObject
    }

    try {
        return Get-MsalToken @connectionParameters
    }
    catch {
        Write-Error $_.Exception.Message
        throw $_
    }
}

function Get-UserGroupMembership {
    <#
    .SYNOPSIS
    Gets the group membership for a given user. Returns an array of objects containing the group name and group id.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $UserEmail
    )

    $accessToken = Get-GraphToken
    $encodedUserEmail = [System.Web.HttpUtility]::UrlEncode($UserEmail)

    try {
        $groupMemberShipResponse = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$encodedUserEmail/memberOf" -Method GET -Headers @{
            Authorization = "Bearer $($accessToken.AccessToken)"
        } | ConvertFrom-Json

        # If @odata.nextLink exists, get next page of results
        while ($groupMemberShipResponse.'@odata.nextLink') {
            $appendGroupMembershipResponse = Invoke-WebRequest -Uri $groupMemberShipResponse.'@odata.nextLink' -Method GET -Headers @{
                Authorization = "Bearer $($accessToken.AccessToken)"
            }
            $graphGroupMembership.value += $appendGroupMembershipResponse.value
            $graphGroupMembership.'@odata.nextLink' = $appendGroupMembershipResponse.'@odata.nextLink'
        }
    }
    catch {
        Write-Error $_.Exception.Message
        throw $_
    }

    $groupMembership = @()
    foreach ($group in $groupMemberShipResponse.value) {
        $groupMembership += [PSCustomObject]@{
            GroupName = $group.displayName
            GroupId   = $group.id
        }
    }

    return $groupMembership
}

function New-CsvFile {
    <#
    .SYNOPSIS
    Creates a new CSV file.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $csv = [PSCustomObject]@{
        UserPrincipalName = $null
        SiteUrl           = $null
        SiteAdmin         = $null
        GroupName         = $null
        PermissionLevel   = $null
        ListName          = $null
        ListPermission    = $null
        TotalRuntimeSeconds = $null
    }

    if (Test-Path $Path) {
        Remove-Item $Path
    }

    $csv | Export-Csv -Path $Path -NoTypeInformation

    # Remove the first (empty) line of the CSV file
    $csvFile = Get-Content $Path
    $csvFile = $csvFile[0..($csvFile.Length - 2)]
    Set-Content -Path $Path -Value $csvFile
}

function Test-UserIsSiteCollectionAdmin {
    <#
    .SYNOPSIS
    Checks if a given user is a site collection admin for a given site collection.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $UserEmail,
        [Parameter(Mandatory = $false)]
        [array] $GraphGroups
    )

    $siteAdmins = Get-PnPSiteCollectionAdmin
    foreach ($siteAdmin in $siteAdmins) {
        $siteAdminLogin = $siteAdmin.LoginName.Split('|')[2]

        if ($UserEmail -eq $siteAdminLogin) {
            return $true
        }

        # Check if user is a member of a group that is a site collection admin
        if ($null -ne $GraphGroups) {
            if ($userGroupMembership.GroupId -contains $siteAdminLogin) {
                return $true
            }
        }
    }

    return $false
}

function Get-UserSharePointGroups {
    <#
    .SYNOPSIS
    Returns an array of SharePoint groups that a given user is a member of for a given site collection.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $UserEmail,
        [Parameter(Mandatory = $false)]
        [array] $GraphGroups
    )

    $siteGroups = Get-PnPGroup

    $groupMembership = @()
    foreach ($siteGroup in $siteGroups) {
        $groupMembers = Get-PnPGroupMember -Identity $siteGroup.Title

        foreach ($groupMember in $groupMembers) {
            $groupMemberLogin = $groupMember.LoginName.Split('|')[2]
            if ($UserEmail -eq $groupMemberLogin) {
                $groupPermissionLevel = Get-PnPGroupPermissions -Identity $siteGroup
                $permissionLevelString = ""
                foreach ($permissionLevel in $groupPermissionLevel) {
                    $permissionLevelString += $permissionLevel.Name + " | "
                }

                if ($permissionLevelString -eq "") {
                    $permissionLevelString = "No Permissions"
                }
                else {
                    # remove trailing " | "
                    $permissionLevelString = $permissionLevelString.Substring(0, $permissionLevelString.Length - 3)
                }

                $groupMembership += [PSCustomObject]@{
                    GroupName       = $siteGroup.Title
                    PermissionLevel = $permissionLevelString
                }

            }
            elseif ($null -ne $GraphGroups) {
                if ($userGroupMembership.GroupId -contains $groupMemberLogin) {
                    $groupPermissionLevel = Get-PnPGroupPermissions -Identity $siteGroup
                    $permissionLevelString = ""
                    foreach ($permissionLevel in $groupPermissionLevel) {
                        $permissionLevelString += $permissionLevel.Name + " | "
                    }

                    if ($permissionLevelString -eq "") {
                        $permissionLevelString = "No Permissions"
                    }
                    else {
                        # remove trailing " | "
                        $permissionLevelString = $permissionLevelString.Substring(0, $permissionLevelString.Length - 3)
                    }

                    $groupMembership += [PSCustomObject]@{
                        GroupName       = $siteGroup.Title
                        PermissionLevel = $permissionLevelString
                    }
                }
            }
        }
    }

    return $groupMembership
}

function Get-UniqueListPermissions {
    <#
    .SYNOPSIS
    Gets the unique permissions at the list level for a given user for a given site collection.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $UserEmail,
        [Parameter(Mandatory = $false)]
        [array] $GraphGroups
    )

    $ctx = Get-PnPContext
    $web = $ctx.Web
    $ctx.Load($web)
    $ctx.ExecuteQuery()

    $lists = $web.Lists
    $ctx.Load($lists)
    $ctx.ExecuteQuery()

    # Exlude built-in lists
    $excludedLists = @("App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms", "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Style Library", , "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", "Reusable Content", "Search Config List", "Solution Gallery", "Site Collection Images", "Suggested Content Browser Locations", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", "Workflow History", "Workflow Tasks", "Preservation Hold Library", "SharePointHomeCacheList")
    $siteListPermissions = @()
    foreach ($list in $lists) {
        $ctx.Load($list)
        $ctx.ExecuteQuery()

        if ($excludedLists -contains $list.Title) {
            continue
        }

        $list.Retrieve("HasUniqueRoleAssignments")
        $ctx.ExecuteQuery()

        if ($list.HasUniqueRoleAssignments) {
            $listPermissions = $list.RoleAssignments
            $ctx.Load($listPermissions)
            $ctx.ExecuteQuery()

            foreach ($roleassignment in $listPermissions) {
                $ctx.Load($roleassignment.Member)
                $ctx.Load($roleassignment.RoleDefinitionBindings)
                $ctx.ExecuteQuery()

                if ($UserEmail -eq ($roleassignment.Member.LoginName.Split('|')[2])) {
                    $listPermission = [PSCustomObject]@{
                        Name            = $list.Title
                        PermissionLevel = $roleassignment.RoleDefinitionBindings.Name
                    }

                    $siteListPermissions += $listPermission
                }
                elseif ($null -ne $GraphGroups) {
                    if ( $GraphGroups.GroupId -contains ($roleassignment.Member.LoginName.Split('|')[2])) {
                        $listPermission = [PSCustomObject]@{
                            Name            = $list.Title
                            PermissionLevel = $roleassignment.RoleDefinitionBindings.Name
                        }

                        $siteListPermissions += $listPermission
                    }
                }
            }
        }
    }
    return $siteListPermissions
}

Set-Location $PSScriptRoot

# Initialize optional transcript logging
$transcriptStarted = $false
if ($PSBoundParameters.ContainsKey('Log') -and $Log) {
    try {
        $logDirectory = Split-Path -Path $Log -Parent
        if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }

        if ((Test-Path -LiteralPath $Log) -and -not $AppendLog) {
            Remove-Item -LiteralPath $Log -Force -ErrorAction SilentlyContinue
        }

        Start-Transcript -Path $Log -Append:$AppendLog -Force | Out-Null
        $transcriptStarted = $true
    }
    catch {
        Write-Warning "Failed to start transcript at '$Log': $($_.Exception.Message)"
    }
}

Write-Host "$(Get-Date) INFO: Connecting to tenant admin site..."
Connect-TenantSite -SiteUrl "https://$TenantName-admin.sharepoint.com" -ErrorAction Stop

Write-Host "$(Get-Date) INFO: Getting all site collections..."
$siteCollections = Get-PnPTenantSite -ErrorAction Stop
Write-Host "$(Get-Date) INFO: `tFound $($siteCollections.Count) site collections."
Disconnect-PnPOnline

Write-Host "$(Get-Date) INFO: Getting group membership for $UserEmail..."
$userGroupMembership = Get-UserGroupMembership -UserEmail $UserEmail -ErrorAction Stop
Write-Host "$(Get-Date) INFO: `tFound $($userGroupMembership.Count) groups."

if (!$Append) {
    New-CsvFile -Path $CSVPath
}

# Prepare inputs for parallel processing
$excludedListsForParallel = @(
    "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", "Composed Looks",
    "Content and Structure Reports", "Content type publishing error log", "Converted Forms", "Device Channels",
    "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", "List Template Gallery",
    "Long Running Operation Status", "Maintenance Log Library", "Style Library", , "Master Docs", "Master Page Gallery",
    "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", "Reusable Content", "Search Config List",
    "Solution Gallery", "Site Collection Images", "Suggested Content Browser Locations", "TaxonomyHiddenList",
    "User Information List", "Web Part Gallery", "wfpub", "wfsvc", "Workflow History", "Workflow Tasks",
    "Preservation Hold Library", "SharePointHomeCacheList"
)

# Convert certificate password once for runspaces
$certPasswordPlain = $null
if ($CertificatePassword) {
    $certPasswordPlain = [System.Net.NetworkCredential]::new('', $CertificatePassword).Password
}

# Read certificate bytes once to avoid concurrent file handle issues in runspaces
$certBytes = [System.IO.File]::ReadAllBytes($CertificatePath)
$certBase64 = [System.Convert]::ToBase64String($certBytes)

Write-Host "$(Get-Date) INFO: Processing $($siteCollections.Count) sites in parallel..."

<# no SPO token required; each runspace authenticates with cert #>

$graphGroupIds = $userGroupMembership.GroupId

$allSiteRows = $siteCollections | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
    $site = $PSItem

    try {
        Import-Module PnP.PowerShell -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Error "Failed to import PnP.PowerShell in parallel block: $($_.Exception.Message)"
        return
    }

    function Invoke-WithRetry {
        param(
            [Parameter(Mandatory = $true)] [scriptblock] $Action,
            [int] $MaxAttempts = 6,
            [int] $InitialDelayMs = 500,
            [string] $OperationName = '',
            [string] $SiteUrlForLog = ''
        )

        $attempt = 0
        while ($true) {
            try {
                return & $Action
            }
            catch {
                $attempt++
                $message = $_.Exception.Message
                $isThrottled = (
                    $message -match '429' -or
                    $message -match 'Too Many Requests' -or
                    $message -match '503' -or
                    $message -match 'Service Unavailable' -or
                    $message -match 'not signed in' -or
                    $message -match 'm_safeCertContext is an invalid handle'
                )
                if (-not $isThrottled -or $attempt -ge $MaxAttempts) {
                    throw
                }

                $delay = [math]::Min(30000, $InitialDelayMs * [math]::Pow(2, $attempt - 1))
                $jitter = Get-Random -Minimum ($delay * 0.9) -Maximum ($delay * 1.1)
                Write-Warning ("{0} THROTTLED: site={1} op={2} attempt={3} delayMs={4} msg={5}" -f (Get-Date), $SiteUrlForLog, $OperationName, $attempt, [int]$jitter, $message)
                Start-Sleep -Milliseconds [int]$jitter
            }
        }
    }

    # Connect per-site using cert auth (base64-encoded PFX)
    try {
        $connection = Invoke-WithRetry -Action {
            if ($using:certPasswordPlain) {
                Connect-PnPOnline -Url $site.Url -ClientId $using:ClientId -CertificateBase64Encoded $using:certBase64 -CertificatePassword (ConvertTo-SecureString $using:certPasswordPlain -AsPlainText -Force) -Tenant "$($using:TenantName).onmicrosoft.com" -ReturnConnection
            }
            else {
                Connect-PnPOnline -Url $site.Url -ClientId $using:ClientId -CertificateBase64Encoded $using:certBase64 -Tenant "$($using:TenantName).onmicrosoft.com" -ReturnConnection
            }
        } -OperationName "Connect" -SiteUrlForLog $site.Url
    }
    catch {
        Write-Warning "Failed to connect to $($site.Url): $($_.Exception.Message)"
        return
    }

    $rows = @()

    try {
        # 1) Site collection admin check
        $isSiteAdmin = $false
        $siteAdmins = Invoke-WithRetry -Action { Get-PnPSiteCollectionAdmin -Connection $connection } -OperationName "Get-PnPSiteCollectionAdmin" -SiteUrlForLog $site.Url
        foreach ($siteAdmin in $siteAdmins) {
            $siteAdminLogin = $siteAdmin.LoginName.Split('|')[2]
            if ($using:UserEmail -eq $siteAdminLogin -or ($using:graphGroupIds -contains $siteAdminLogin)) {
                $isSiteAdmin = $true
                break
            }
        }

        if ($isSiteAdmin) {
            $rows += [PSCustomObject]@{
                UserPrincipalName   = $using:UserEmail
                SiteUrl             = $site.Url
                SiteAdmin           = $true
                GroupName           = $null
                PermissionLevel     = $null
                ListName            = $null
                ListPermission      = $null
                TotalRuntimeSeconds = $null
            }

            return $rows
        }

        # 2) SharePoint group membership check
        $siteGroups = Invoke-WithRetry -Action { Get-PnPGroup -Connection $connection } -OperationName "Get-PnPGroup" -SiteUrlForLog $site.Url
        foreach ($siteGroup in $siteGroups) {
            try {
                $groupMembers = Invoke-WithRetry -Action { Get-PnPGroupMember -Identity $siteGroup.Title -Connection $connection } -OperationName ("Get-PnPGroupMember:{0}" -f $siteGroup.Title) -SiteUrlForLog $site.Url
            }
            catch {
                if ($_.Exception.Message -match 'Group cannot be found') { continue }
                throw
            }
            foreach ($groupMember in $groupMembers) {
                $groupMemberLogin = $groupMember.LoginName.Split('|')[2]
                if ($using:UserEmail -eq $groupMemberLogin -or ($using:graphGroupIds -contains $groupMemberLogin)) {
                    $groupPermissionLevel = Invoke-WithRetry -Action { Get-PnPGroupPermissions -Identity $siteGroup -Connection $connection } -OperationName ("Get-PnPGroupPermissions:{0}" -f $siteGroup.Title) -SiteUrlForLog $site.Url
                    $permissionLevelString = if ($groupPermissionLevel) { ($groupPermissionLevel | ForEach-Object { $_.Name }) -join ' | ' } else { 'No Permissions' }

                    $rows += [PSCustomObject]@{
                        UserPrincipalName   = $using:UserEmail
                        SiteUrl             = $site.Url
                        SiteAdmin           = $false
                        GroupName           = $siteGroup.Title
                        PermissionLevel     = $permissionLevelString
                        ListName            = $null
                        ListPermission      = $null
                        TotalRuntimeSeconds = $null
                    }

                    break
                }
            }
        }

        # 3) Unique list-level permissions
        $ctx = Get-PnPContext -Connection $connection
        $web = $ctx.Web
        $ctx.Load($web)
        Invoke-WithRetry -Action { $ctx.ExecuteQuery() } -OperationName "CSOM:LoadWeb" -SiteUrlForLog $site.Url | Out-Null

        $lists = $web.Lists
        $ctx.Load($lists)
        Invoke-WithRetry -Action { $ctx.ExecuteQuery() } -OperationName "CSOM:LoadLists" -SiteUrlForLog $site.Url | Out-Null

        foreach ($list in $lists) {
            $ctx.Load($list)
            Invoke-WithRetry -Action { $ctx.ExecuteQuery() } -OperationName ("CSOM:LoadList:{0}" -f $list.Title) -SiteUrlForLog $site.Url | Out-Null

            if ($using:excludedListsForParallel -contains $list.Title) { continue }

            $list.Retrieve("HasUniqueRoleAssignments")
            Invoke-WithRetry -Action { $ctx.ExecuteQuery() } -OperationName ("CSOM:HasUniqueRoleAssignments:{0}" -f $list.Title) -SiteUrlForLog $site.Url | Out-Null

            if ($list.HasUniqueRoleAssignments) {
                $listPermissions = $list.RoleAssignments
                $ctx.Load($listPermissions)
                Invoke-WithRetry -Action { $ctx.ExecuteQuery() } -OperationName ("CSOM:LoadRoleAssignments:{0}" -f $list.Title) -SiteUrlForLog $site.Url | Out-Null

                foreach ($roleassignment in $listPermissions) {
                    $ctx.Load($roleassignment.Member)
                    $ctx.Load($roleassignment.RoleDefinitionBindings)
                    Invoke-WithRetry -Action { $ctx.ExecuteQuery() } -OperationName ("CSOM:LoadAssignment:{0}" -f $list.Title) -SiteUrlForLog $site.Url | Out-Null

                    $loginToCheck = $roleassignment.Member.LoginName.Split('|')[2]
                    if ($using:UserEmail -eq $loginToCheck -or ($using:graphGroupIds -contains $loginToCheck)) {
                        $rows += [PSCustomObject]@{
                            UserPrincipalName   = $using:UserEmail
                            SiteUrl             = $site.Url
                            SiteAdmin           = $false
                            GroupName           = $null
                            PermissionLevel     = $null
                            ListName            = $list.Title
                            ListPermission      = $roleassignment.RoleDefinitionBindings.Name
                            TotalRuntimeSeconds = $null
                        }
                        break
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Failed processing $($site.Url): $($_.Exception.Message)"
    }
    finally {
        try { Disconnect-PnPOnline -Connection $connection | Out-Null } catch { }
    }

    return $rows

}

# Flatten results and write once
$rowsToWrite = @()
if ($allSiteRows) { $rowsToWrite += $allSiteRows }
if ($rowsToWrite.Count -gt 0) {
    $rowsToWrite | Export-Csv -Path $CSVPath -Append -NoTypeInformation
}

# Append total runtime summary row for this user
$scriptEndTime = Get-Date
$elapsed = $scriptEndTime - $scriptStartTime
$totalSeconds = [math]::Round($elapsed.TotalSeconds, 2)
Write-Host "$(Get-Date) INFO: Total runtime for $($UserEmail): $($totalSeconds) seconds."

$csvLineObject = [PSCustomObject]@{
    UserPrincipalName   = $UserEmail
    SiteUrl             = $null
    SiteAdmin           = $null
    GroupName           = $null
    PermissionLevel     = $null
    ListName            = $null
    ListPermission      = $null
    TotalRuntimeSeconds = $totalSeconds
}
$csvLineObject | Export-Csv -Path $CSVPath -Append -NoTypeInformation

if ($transcriptStarted) {
    try { Stop-Transcript | Out-Null } catch { }
}
