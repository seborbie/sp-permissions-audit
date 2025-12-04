# Get-SharePointTenantPermissions.ps1
# Description: This script will get all the permissions for a given user or users in a SharePoint Online tenant and export them to a CSV file.

#requires -Version 7.5.4
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
    [securestring] $CertificatePassword,
    [Parameter(Mandatory = $false)]
    [switch] $Append = $false,
    [Parameter(Mandatory = $false)]
    [string] $Log,
    [Parameter(Mandatory = $false)]
    [switch] $AppendLog,
    [Parameter(Mandatory = $false)]
    [int] $ThrottleLimit = 1,  # Number of parallel threads (increase to enable concurrent site processing)
    [Parameter(Mandatory = $false)]
    [int] $Max406Retries = 999,
    [Parameter(Mandatory = $false)]
    [switch] $ResolveItemSharingLinks,
    [Parameter(Mandatory = $false)]
    [int] $MaxItemScanPerList = 500
)

# Ensure MSAL.PS is available for the current user (non-interactive install on first run)
try {
    # Prefer TLS 1.2 for gallery operations when needed
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

    if (-not (Get-Module -ListAvailable -Name 'MSAL.PS')) {
        try { $null = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction Stop } catch { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null }
        $psg = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
        if (-not $psg) { Register-PSRepository -Default -ErrorAction Stop }
        $psg = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
        if ($psg -and $psg.InstallationPolicy -ne 'Trusted') { Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted }
        Install-Module -Name 'MSAL.PS' -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module 'MSAL.PS' -ErrorAction Stop
}
catch {
    Write-Error "Failed to ensure MSAL.PS module is installed/imported: $($_.Exception.Message)"
    throw
}

# Start benchmarking for this user
$scriptStartTime = Get-Date

# Console verbosity control: when -Log is supplied, keep console output minimal
$ConsoleQuiet = $false
if ($PSBoundParameters.ContainsKey('Log') -and $Log) { $ConsoleQuiet = $true }

# Defaults for new sharing link resolution (enabled unless explicitly disabled)
if (-not $PSBoundParameters.ContainsKey('ResolveItemSharingLinks')) { $ResolveItemSharingLinks = $true }
if (-not $PSBoundParameters.ContainsKey('MaxItemScanPerList') -or $MaxItemScanPerList -le 0) { $MaxItemScanPerList = 200 }

# Buffer log lines when console is quiet to avoid file locking during execution
$LogBuffer = $null
if ($ConsoleQuiet -and $Log) {
    $LogBuffer = New-Object System.Collections.ArrayList
}

function Write-Detail {
    param(
        [Parameter(Mandatory = $true)] [string] $Message
    )
    if ($ConsoleQuiet -and $Log) { [void]$LogBuffer.Add($Message) } else { Write-Host $Message }
}

function Write-Major {
    param(
        [Parameter(Mandatory = $true)] [string] $Message,
        [Parameter(Mandatory = $false)] [System.ConsoleColor] $ForegroundColor
    )
    # Always show on console
    if ($PSBoundParameters.ContainsKey('ForegroundColor') -and $null -ne $ForegroundColor) { Write-Host $Message -ForegroundColor $ForegroundColor } else { Write-Host $Message }
    # Also capture in log when quiet mode is on
    if ($ConsoleQuiet -and $Log) { [void]$LogBuffer.Add($Message) }
}

function Write-Warn {
    param(
        [Parameter(Mandatory = $true)] [string] $Message
    )
    if ($ConsoleQuiet -and $Log) {
        [void]$LogBuffer.Add("WARNING: $Message")
    } else {
        Write-Warning $Message
    }
}

 

function Get-GraphToken {
    <#
    .SYNOPSIS
    Gets a bearer token for the Microsoft Graph API using certificate-based authentication.
    #>
    # Load the certificate, using the password if provided
    if ($CertificatePassword) {
        $passwordBstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePassword)
        try {
            $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordBstr)
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr)
        }
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $plainPassword)
    }
    else {
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
    }

    $connectionParameters = @{
        'TenantId'          = "$TenantName.onmicrosoft.com"
        'ClientId'          = $ClientId
        'ClientCertificate' = $certificate
    }

    try {
        return Get-MsalToken @connectionParameters
    }
    catch {
        Write-Error $_.Exception.Message
        throw $_
    }
}

function Get-GraphUserObjectId {
	param(
		[Parameter(Mandatory = $true)] [string] $UserEmail
	)
	$token = Get-GraphToken
	$encodedUserEmail = [System.Web.HttpUtility]::UrlEncode($UserEmail)
	try {
		$user = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$encodedUserEmail?$select=id,userPrincipalName,mail" -Method GET -Headers @{ Authorization = "Bearer $($token.AccessToken)" } | ConvertFrom-Json
		return ($user.id)
	} catch { return $null }
}

function Get-GraphSiteId {
	param(
		[Parameter(Mandatory = $true)] [string] $SiteUrl
	)
	$token = Get-GraphToken
	try {
		$u = [Uri]$SiteUrl
		$host = $u.Host
		$path = $u.AbsolutePath
		if ([string]::IsNullOrWhiteSpace($path)) { $path = '/' }
		$resp = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($host):$($path)?$select=id" -Method GET -Headers @{ Authorization = "Bearer $($token.AccessToken)" } -ErrorAction Stop | ConvertFrom-Json
		return $resp.id
	} catch { return $null }
}

function Get-GraphListItemPermissionsByIds {
	param(
		[Parameter(Mandatory = $true)] [string] $SiteUrl,
		[Parameter(Mandatory = $true)] [string] $ListId,
		[Parameter(Mandatory = $true)] [int] $ItemId,
		[Parameter(Mandatory = $true)] [string] $GraphAccessToken
	)

	try {
		$siteId = Get-GraphSiteId -SiteUrl $SiteUrl
		if (-not $siteId) { return $null }
		$item = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$ListId/items/$ItemId/driveItem?$select=id,webUrl" -Method GET -Headers @{ Authorization = "Bearer $GraphAccessToken" } -ErrorAction Stop | ConvertFrom-Json
		if (-not $item -or -not $item.id) { return $null }
		$perms = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$ListId/items/$ItemId/driveItem/permissions" -Method GET -Headers @{ Authorization = "Bearer $GraphAccessToken" } -ErrorAction Stop | ConvertFrom-Json
		return @{ ItemWebUrl = $item.webUrl; Permissions = $perms.value }
	} catch { return $null }
}

function Get-UserAccessibleSharingLinksForListItem {
	param(
		[Parameter(Mandatory = $true)] [string] $SiteUrl,
		[Parameter(Mandatory = $true)] [string] $ListId,
		[Parameter(Mandatory = $true)] [int] $ItemId,
		[Parameter(Mandatory = $true)] [string] $UserEmail,
		[Parameter(Mandatory = $true)] [object] $UserGroups
	)

	$graphToken = Get-GraphToken
	$userId = Get-GraphUserObjectId -UserEmail $UserEmail
	if (-not $graphToken -or -not $graphToken.AccessToken) { return @() }
	$resp = Get-GraphListItemPermissionsByIds -SiteUrl $SiteUrl -ListId $ListId -ItemId $ItemId -GraphAccessToken $graphToken.AccessToken
	if (-not $resp -or -not $resp.Permissions) { return @() }
	$itemUrl = $resp.ItemWebUrl
	$rows = @()
	foreach ($p in $resp.Permissions) {
		if (-not $p.link) { continue }
		$scope = $p.link.scope
		$roles = @(); foreach ($r in $p.roles) { $roles += [string]$r }
		$roleStr = ($roles -join ' | ')
		$accessible = $false
		if ($scope -eq 'anonymous' -or $scope -eq 'organization') { $accessible = $true }
		elseif ($scope -eq 'users') {
			try {
				if ($p.grantedToIdentitiesV2) {
					foreach ($gi in $p.grantedToIdentitiesV2) {
						if ($gi.user -and $userId -and ($gi.user.id -eq $userId)) { $accessible = $true; break }
						if ($gi.group -and $UserGroups -and $UserGroups.GroupId -contains $gi.group.id) { $accessible = $true; break }
					}
				}
			} catch { }
		}
		if ($accessible) {
			$rows += [PSCustomObject]@{
				ItemUrl = $itemUrl
				Permission = $roleStr
				Scope = $scope
			}
		}
	}
	return $rows
}
function Get-GraphDriveItemPermissions {
	param(
		[Parameter(Mandatory = $true)] [string] $SiteUrl,
		[Parameter(Mandatory = $true)] [string] $ServerRelativeFileRef,
		[Parameter(Mandatory = $true)] [string] $GraphAccessToken
	)

	try {
		$siteUri = [Uri]$SiteUrl
		$siteHost = $siteUri.Host
		$sitePath = $siteUri.AbsolutePath
		if ([string]::IsNullOrWhiteSpace($sitePath)) { $sitePath = '/' }
		# Resolve Graph site id
		$siteIdResp = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($siteHost):$($sitePath)?$select=id,webUrl" -Method GET -Headers @{ Authorization = "Bearer $GraphAccessToken" } -ErrorAction Stop | ConvertFrom-Json
		$siteId = $siteIdResp.id
		if (-not $siteId) { return $null }
		# Compute drive and item paths
		$sr = $ServerRelativeFileRef
		# Ensure server relative begins with sitePath
		if (-not $sr.StartsWith($sitePath, [StringComparison]::OrdinalIgnoreCase)) { return $null }
		$afterSite = $sr.Substring($sitePath.Length).TrimStart('/')
		$firstSlash = $afterSite.IndexOf('/')
		if ($firstSlash -lt 0) { $libraryName = $afterSite; $restPath = '' } else { $libraryName = $afterSite.Substring(0, $firstSlash); $restPath = $afterSite.Substring($firstSlash + 1) }
		# Find the drive matching the library
		$drives = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives?$select=id,name,webUrl" -Method GET -Headers @{ Authorization = "Bearer $GraphAccessToken" } -ErrorAction Stop | ConvertFrom-Json
		$drive = $null
		foreach ($d in $drives.value) { if ($d.name -eq $libraryName -or ($d.webUrl -and $d.webUrl -match [Regex]::Escape($libraryName))) { $drive = $d; break } }
		if (-not $drive) { return $null }
		# Resolve item id
		if ([string]::IsNullOrWhiteSpace($restPath)) {
			$item = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root?$select=id,webUrl" -Method GET -Headers @{ Authorization = "Bearer $GraphAccessToken" } -ErrorAction Stop | ConvertFrom-Json
		} else {
			$encodedPath = [System.Web.HttpUtility]::UrlPathEncode($restPath)
			$item = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root:/$([uri]::EscapeDataString($restPath))?$select=id,webUrl" -Method GET -Headers @{ Authorization = "Bearer $GraphAccessToken" } -ErrorAction Stop | ConvertFrom-Json
		}
		if (-not $item -or -not $item.id) { return $null }
		# Get permissions
		$perms = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/items/$($item.id)/permissions" -Method GET -Headers @{ Authorization = "Bearer $GraphAccessToken" } -ErrorAction Stop | ConvertFrom-Json
		return @{ ItemWebUrl = $item.webUrl; Permissions = $perms.value }
	} catch { return $null }
}

function Get-UserAccessibleSharingLinksForItem {
	param(
		[Parameter(Mandatory = $true)] [string] $SiteUrl,
		[Parameter(Mandatory = $true)] [string] $ServerRelativeFileRef,
		[Parameter(Mandatory = $true)] [string] $UserEmail,
		[Parameter(Mandatory = $true)] [object] $UserGroups
	)

	$graphToken = Get-GraphToken
	$userId = Get-GraphUserObjectId -UserEmail $UserEmail
	if (-not $graphToken -or -not $graphToken.AccessToken) { return @() }
	$resp = Get-GraphDriveItemPermissions -SiteUrl $SiteUrl -ServerRelativeFileRef $ServerRelativeFileRef -GraphAccessToken $graphToken.AccessToken
	if (-not $resp -or -not $resp.Permissions) { return @() }
	$itemUrl = $resp.ItemWebUrl
	$rows = @()
	foreach ($p in $resp.Permissions) {
		if (-not $p.link) { continue }
		$scope = $p.link.scope
		$roles = @(); foreach ($r in $p.roles) { $roles += [string]$r }
		$roleStr = ($roles -join ' | ')
		$accessible = $false
		if ($scope -eq 'anonymous' -or $scope -eq 'organization') { $accessible = $true }
		elseif ($scope -eq 'users') {
			# Check if granted to this user or any of their AAD groups
			try {
				if ($p.grantedToIdentitiesV2) {
					foreach ($gi in $p.grantedToIdentitiesV2) {
						if ($gi.user -and $userId -and ($gi.user.id -eq $userId)) { $accessible = $true; break }
						if ($gi.group -and $UserGroups -and $UserGroups.GroupId -contains $gi.group.id) { $accessible = $true; break }
					}
				}
			} catch { }
		}
		if ($accessible) {
			$rows += [PSCustomObject]@{
				ItemUrl = $itemUrl
				Permission = $roleStr
				Scope = $scope
			}
		}
	}
	return $rows
}

function Get-SharePointAccessToken {
    <#
    .SYNOPSIS
    Gets a bearer token for SharePoint REST using certificate-based authentication.
    #>
    if ($CertificatePassword) {
        $passwordBstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePassword)
        try {
            $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordBstr)
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr)
        }
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $plainPassword)
    }
    else {
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
    }

    $connectionParameters = @{
        'TenantId'          = "$TenantName.onmicrosoft.com"
        'ClientId'          = $ClientId
        'ClientCertificate' = $certificate
        'Scopes'            = "https://$TenantName.sharepoint.com/.default"
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
            } | ConvertFrom-Json
            $groupMemberShipResponse.value += $appendGroupMembershipResponse.value
            $groupMemberShipResponse.'@odata.nextLink' = $appendGroupMembershipResponse.'@odata.nextLink'
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
        ItemUrl           = $null
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

function Invoke-SharePointRestWithAcceptFallback {
    param (
        [Parameter(Mandatory = $true)]
		[string] $Uri,
		[Parameter(Mandatory = $true)]
		[hashtable] $BaseHeaders,
		[Parameter(Mandatory = $false)]
		[string] $Method = 'GET',
        [Parameter(Mandatory = $false)]
		[object] $Body
	)

	# Start with pinned nometadata for highest performance (except Search)
	$headers = @{}
	foreach ($key in $BaseHeaders.Keys) { if ($key -ne 'Accept') { $headers[$key] = $BaseHeaders[$key] } }

	$uriWithFormat = $Uri
	$isSearch = $Uri -match '/_api/search/'
	if ($isSearch) {
		# Search API: do NOT append $format. Retry 406s up to Max406Retries rotating Accept variants
		$accepts = @('application/json', 'application/json;odata=minimalmetadata', 'application/json;odata=verbose', '*/*', '')
		$attempt = 0
		while ($attempt -le $Max406Retries) {
			$accept = $accepts[$attempt % $accepts.Count]
			$variantHeaders = @{}
			foreach ($k in $headers.Keys) { if ($k -ne 'Accept') { $variantHeaders[$k] = $headers[$k] } }
			if ([string]::IsNullOrEmpty($accept)) { if ($variantHeaders.ContainsKey('Accept')) { $variantHeaders.Remove('Accept') } } else { $variantHeaders['Accept'] = $accept }
			try {
				if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
					return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -Body $Body -ErrorAction Stop
				} else {
					return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -ErrorAction Stop
				}
			}
			catch {
				$resp = $_.Exception.Response
				$status = $null
				if ($resp -and $resp.StatusCode) { $status = [int]$resp.StatusCode.value__ }
				$shouldNext = ($status -eq 406) -or ($_.Exception.Message -match '406|Not Acceptable')
				if (-not $shouldNext) { throw }
				$attempt++
				continue
			}
		}
		throw "Received 406 Not Acceptable from $Uri after $Max406Retries retries."
	}

	# Non-Search: use OData nometadata and ensure $format matches
	$uriBase = $Uri
	$attempt = 0
	while ($attempt -le $Max406Retries) {
		# Variant 1: nometadata
		$variantHeaders = @{}
		foreach ($k in $headers.Keys) { if ($k -ne 'Accept') { $variantHeaders[$k] = $headers[$k] } }
		$variantHeaders['Accept'] = 'application/json;odata=nometadata'
		$uriWithFormat = $uriBase
		if ($uriWithFormat -match '(\?|&)`?\$format=') {
			$uriWithFormat = $uriWithFormat -replace '(\?|&)`?\$format=[^&]+', '$1`$format=application/json;odata=nometadata'
		} else {
			$separator = ($uriWithFormat -match '\?') ? '&' : '?'
			$uriWithFormat = "$uriWithFormat${separator}`$format=application/json;odata=nometadata"
		}
		try {
			if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
				return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -Body $Body -ErrorAction Stop
			} else {
				return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -ErrorAction Stop
			}
		}
		catch {
			$resp = $_.Exception.Response
			$status = $null
			if ($resp -and $resp.StatusCode) { $status = [int]$resp.StatusCode.value__ }
			$shouldNext = ($status -eq 406) -or ($_.Exception.Message -match '406|Not Acceptable')
			if (-not $shouldNext) { throw }
		}

		# Variant 2: minimalmetadata
		$variantHeaders = @{}
		foreach ($k in $headers.Keys) { if ($k -ne 'Accept') { $variantHeaders[$k] = $headers[$k] } }
		$variantHeaders['Accept'] = 'application/json;odata=minimalmetadata'
		$uriWithFormat = $uriBase
		if ($uriWithFormat -match '(\?|&)`?\$format=') {
			$uriWithFormat = $uriWithFormat -replace '(\?|&)`?\$format=[^&]+', '$1`$format=application/json;odata=minimalmetadata'
		} else {
			$separator = ($uriWithFormat -match '\?') ? '&' : '?'
			$uriWithFormat = "$uriWithFormat${separator}`$format=application/json;odata=minimalmetadata"
		}
		try {
			if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
				return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -Body $Body -ErrorAction Stop
			} else {
				return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -ErrorAction Stop
			}
		}
		catch {
			$resp = $_.Exception.Response
			$status = $null
			if ($resp -and $resp.StatusCode) { $status = [int]$resp.StatusCode.value__ }
			$shouldNext = ($status -eq 406) -or ($_.Exception.Message -match '406|Not Acceptable')
			if (-not $shouldNext) { throw }
		}

		$attempt++
	}

	throw "Received 406 Not Acceptable from $Uri after $Max406Retries retries."
}

function Resolve-SharingLinkPrincipalToItemUrl {
	param(
		[Parameter(Mandatory = $true)] [string] $SiteUrl,
		[Parameter(Mandatory = $true)] [string] $AccessToken,
		[Parameter(Mandatory = $true)] [string] $PrincipalTitle
	)

	# Extract GUID from principal title (e.g., SharingLinks.<GUID>)
	$guidMatch = [regex]::Match($PrincipalTitle, '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})')
	if (-not $guidMatch.Success) { return $null }
	$linkGuid = $guidMatch.Groups[1].Value

	$headers = @{ Authorization = "Bearer $AccessToken"; Accept = 'application/json;odata=nometadata' }
	try {
		$fileResp = Invoke-SharePointRestWithAcceptFallback -Uri ("$SiteUrl/_api/web/GetFileById(guid'$linkGuid')") -BaseHeaders $headers -Method GET
		if ($fileResp -and $fileResp.ServerRelativeUrl) {
			$su = [Uri]$SiteUrl
			return "$($su.Scheme)://$($su.Host)$($fileResp.ServerRelativeUrl)"
		}
	} catch { }
	try {
		$folderResp = Invoke-SharePointRestWithAcceptFallback -Uri ("$SiteUrl/_api/web/GetFolderById(guid'$linkGuid')") -BaseHeaders $headers -Method GET
		if ($folderResp -and $folderResp.ServerRelativeUrl) {
			$su = [Uri]$SiteUrl
			return "$($su.Scheme)://$($su.Host)$($folderResp.ServerRelativeUrl)"
		}
	} catch { }

	# Fallback: Search by UniqueId for the path (fast and indexed)
	try {
		$searchUri = "$SiteUrl/_api/search/query?querytext='UniqueId:$linkGuid'&selectproperties='Path'&rowlimit=1"
		$searchResp = Invoke-SharePointRestWithAcceptFallback -Uri $searchUri -BaseHeaders $headers -Method GET
		if ($searchResp -and $searchResp.PrimaryQueryResult -and $searchResp.PrimaryQueryResult.RelevantResults -and $searchResp.PrimaryQueryResult.RelevantResults.Table) {
			foreach ($row in $searchResp.PrimaryQueryResult.RelevantResults.Table.Rows) {
				$props = @{}
				foreach ($cell in $row.Cells) { $props[$cell.Key] = $cell.Value }
				if ($props['Path']) { return [string]$props['Path'] }
			}
		}
	} catch { }
	return $null
}

function Search-ItemsSharedWithUserRest {
	param(
		[Parameter(Mandatory = $true)] [string] $UserEmail,
		[Parameter(Mandatory = $false)] [string] $FilterSiteUrl
	)

	# Build certificate
	if ($CertificatePassword) {
		$passwordBstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePassword)
		try { $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordBstr) } finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
		$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $plainPassword)
	} else {
		$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
	}

	$adminHost = "$TenantName-admin.sharepoint.com"
	$spoToken = Get-MsalToken -TenantId "$TenantName.onmicrosoft.com" -ClientId $ClientId -ClientCertificate $certificate -Scopes "https://$adminHost/.default"
	$headers = @{ Authorization = "Bearer $($spoToken.AccessToken)"; Accept = 'application/json;odata=nometadata' }

	# Construct robust identity token set for search (email, guest UPN, membership claim)
	$userTokens = New-Object System.Collections.Generic.List[string]
	$ue = ($UserEmail.Trim()).ToLowerInvariant()
	if (-not [string]::IsNullOrWhiteSpace($ue)) { [void]$userTokens.Add($ue) }
	try {
		$gt = Get-GraphToken
		if ($gt -and $gt.AccessToken) {
			$encoded = [System.Web.HttpUtility]::UrlEncode($UserEmail)
			$u = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$encoded?$select=userPrincipalName" -Method GET -Headers @{ Authorization = "Bearer $($gt.AccessToken)" } -ErrorAction SilentlyContinue | ConvertFrom-Json
			if ($u -and $u.userPrincipalName) {
				$upnLower = ($u.userPrincipalName).ToLowerInvariant()
				if (-not [string]::IsNullOrWhiteSpace($upnLower)) { [void]$userTokens.Add($upnLower) }
			}
		}
	} catch { }
	# Add membership claim variants
	$claimTokens = New-Object System.Collections.Generic.List[string]
	foreach ($tok in $userTokens) { [void]$claimTokens.Add("i:0#.f|membership|$tok") }

	# Build KQL
	$kqlParts = New-Object System.Collections.Generic.List[string]
    foreach ($tok in $userTokens) { [void]$kqlParts.Add(('SharedWithUsersOWSUSER="' + $tok + '"')) }
    foreach ($ctok in $claimTokens) { [void]$kqlParts.Add(('SharedWithUsersOWSUSER="' + $ctok + '"')) }
	$kqlIdentity = '(' + ([string]::Join(' OR ', $kqlParts)) + ')'
	$kql = "$kqlIdentity AND contentclass:STS_ListItem_DocumentLibrary"
    if ($FilterSiteUrl) {
        # Constrain to current site path to reduce noise
        $kql += (' AND path="' + $FilterSiteUrl + '"')
    }
	$q = [System.Web.HttpUtility]::UrlEncode($kql)

	$startRow = 0
	$rowLimit = 500
	$items = New-Object System.Collections.Generic.List[object]
	$more = $true
	while ($more) {
		$uri = "https://$adminHost/_api/search/query?querytext='$q'&rowlimit=$rowLimit&startrow=$startRow&trimduplicates=false&selectproperties='Path,UniqueId'"
		$response = Invoke-SharePointRestWithAcceptFallback -Uri $uri -BaseHeaders $headers -Method GET
		$results = $response.PrimaryQueryResult.RelevantResults
		if ($results -and $results.Table -and $results.Table.Rows) {
			foreach ($row in $results.Table.Rows) {
				$props = @{}
				foreach ($cell in $row.Cells) { $props[$cell.Key] = $cell.Value }
				if ($props.ContainsKey('Path') -and $props['Path']) {
					$item = [PSCustomObject]@{ ItemUrl = [string]$props['Path'] }
					[void]$items.Add($item)
				}
			}
		}
		$startRow += $rowLimit
		$more = ($results -and $results.TotalRows -gt $startRow)
	}

	return $items
}

function Get-ItemSharingLinksRest {
	param(
		[Parameter(Mandatory = $true)] [string] $ItemUrl,
		[Parameter(Mandatory = $true)] [string] $AccessToken
	)

	try {
		$uri = [Uri]$ItemUrl
		$apiBase = "$($uri.Scheme)://$($uri.Host)/_api/SP.Sharing.DocumentSharingManager.GetObjectSharingInformation"
		$headers = @{ Authorization = "Bearer $AccessToken"; Accept = 'application/json;odata=nometadata'; 'Content-Type' = 'application/json;odata=nometadata' }
		$body = @{ resourceAddress = $ItemUrl; groupId = 0; useSimplifiedRoles = $true; useAppTokenOnly = $true } | ConvertTo-Json -Depth 5
		$resp = Invoke-RestMethod -Method POST -Uri $apiBase -Headers $headers -Body $body -ErrorAction Stop
		return $resp
	} catch { return $null }
}

function Try-ResolveSharingLinkBySamplingItems {
	param(
		[Parameter(Mandatory = $true)] [string] $SiteUrl,
		[Parameter(Mandatory = $true)] [string] $ListId,
		[Parameter(Mandatory = $true)] [string] $PrincipalTitle,
		[Parameter(Mandatory = $true)] [int] $MaxScan,
		[Parameter(Mandatory = $true)] [string] $AccessToken
	)

	$headers = @{ Authorization = "Bearer $AccessToken"; Accept = 'application/json;odata=nometadata' }
	try {
		$endpoint = "$SiteUrl/_api/web/lists(guid'$ListId')/items?`$select=Id,FileRef,HasUniqueRoleAssignments,RoleAssignments/Member/Title&`$filter=HasUniqueRoleAssignments%20eq%20true&`$expand=RoleAssignments,RoleAssignments/Member&`$top=$MaxScan"
		$response = Invoke-SharePointRestWithAcceptFallback -Uri $endpoint -BaseHeaders $headers -Method GET
		$items = @()
		if ($response) { if ($response.value) { $items = $response.value } else { $items = @($response) } }
		foreach ($i in $items) {
			if (-not $i.RoleAssignments) { continue }
			foreach ($ra in $i.RoleAssignments) {
				if ($ra.Member -and $ra.Member.Title -eq $PrincipalTitle) {
					$su = [Uri]$SiteUrl
					if ($i.FileRef) { return "$($su.Scheme)://$($su.Host)$($i.FileRef)" }
				}
			}
		}
	} catch { }
	return $null
}

function Get-UserOneDriveSiteUrl {
    <#
    .SYNOPSIS
	Resolves the target user's OneDrive (personal site) root URL using Microsoft Graph.
    #>
    param (
        [Parameter(Mandatory = $true)]
		[string] $UserEmail
	)

	$graphToken = Get-GraphToken
	$encodedUserEmail = [System.Web.HttpUtility]::UrlEncode($UserEmail)
	try {
		$drive = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$encodedUserEmail/drive" -Method GET -Headers @{ Authorization = "Bearer $($graphToken.AccessToken)" } | ConvertFrom-Json
		$webUrl = $drive.webUrl
		if ([string]::IsNullOrEmpty($webUrl)) { return $null }
		# Typical format: https://{tenant}-my.sharepoint.com/personal/{normalized_upn}/Documents
		$documentsIndex = $webUrl.IndexOf('/Documents', [StringComparison]::OrdinalIgnoreCase)
		if ($documentsIndex -gt 0) {
			return $webUrl.Substring(0, $documentsIndex)
		}
		# Fallback: return parent segment without trailing slash
		return $webUrl.TrimEnd('/')
	}
	catch {
		return $null
	}
}

function Get-TenantSitesRest {
    <#
    .SYNOPSIS
	Enumerates site collections using SharePoint Search REST and filters out other users' personal sites.
    #>
    param (
        [Parameter(Mandatory = $true)]
		[string] $UserEmail
	)

	# Build certificate
	if ($CertificatePassword) {
		$passwordBstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePassword)
		try { $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordBstr) } finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
		$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $plainPassword)
                }
                else {
		$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
	}

	$adminHost = "$TenantName-admin.sharepoint.com"
	$spoToken = Get-MsalToken -TenantId "$TenantName.onmicrosoft.com" -ClientId $ClientId -ClientCertificate $certificate -Scopes "https://$adminHost/.default"
	$headers = @{ Authorization = "Bearer $($spoToken.AccessToken)"; Accept = 'application/json;odata=nometadata' }

	$startRow = 0
	$rowLimit = 500
	$allUrls = New-Object System.Collections.Generic.List[string]

	# Retry entire search until we get non-zero collections or attempts exhausted
	$outerAttempts = 0
	$foundAny = $false
	while (-not $foundAny -and $outerAttempts -le $Max406Retries) {
		$outerAttempts++
		$startRow = 0
		$allUrls.Clear()
		do {
			$uri = "https://$adminHost/_api/search/query?querytext='contentclass:STS_Site'&rowlimit=$rowLimit&startrow=$startRow&trimduplicates=false&selectproperties='Path'"
			$response = Invoke-SharePointRestWithAcceptFallback -Uri $uri -BaseHeaders $headers -Method GET
			$results = $response.PrimaryQueryResult.RelevantResults
			if ($results -and $results.Table -and $results.Table.Rows) {
				foreach ($row in $results.Table.Rows) {
					$props = @{}
					foreach ($cell in $row.Cells) { $props[$cell.Key] = $cell.Value }
					if ($props.ContainsKey('Path') -and $props['Path']) {
						[void]$allUrls.Add($props['Path'])
					}
				}
			}
			$startRow += $rowLimit
		} while ($results -and $results.TotalRows -gt $startRow)

		if ($allUrls.Count -gt 0) { $foundAny = $true }
	}

	# Filter out other users' personal sites; keep only the current user's OneDrive
	$uniqueUrls = $allUrls | Select-Object -Unique
	$tenantRoot = "https://$TenantName.sharepoint.com"
	$myHost = "https://$TenantName-my.sharepoint.com"
	# Only resolve OneDrive if at least one URL is on the -my host
	$userOneDrive = $null
	$hasMyHostUrl = $false
	foreach ($u in $uniqueUrls) { if ($u.StartsWith($myHost, [StringComparison]::OrdinalIgnoreCase)) { $hasMyHostUrl = $true; break } }
	if ($hasMyHostUrl) { $userOneDrive = Get-UserOneDriveSiteUrl -UserEmail $UserEmail }

	$filtered = @()
	foreach ($u in $uniqueUrls) {
		if ($u.StartsWith($tenantRoot, [StringComparison]::OrdinalIgnoreCase)) {
			$filtered += $u
            continue
        }
		if ($u.StartsWith($myHost, [StringComparison]::OrdinalIgnoreCase)) {
			if ($null -ne $userOneDrive -and $u.StartsWith($userOneDrive, [StringComparison]::OrdinalIgnoreCase)) {
				$filtered += $u
			}
			continue
		}
		# Exclude other hosts by default
	}

	# Exclude content storage URLs
	$filtered = $filtered | Where-Object { $_ -notmatch '/contentstorage' }

	return $filtered | ForEach-Object { [PSCustomObject]@{ Url = $_ } }
}

function Is-TransientSharePointError {
    param(
        [Parameter(Mandatory = $true)] [object] $ErrorObject
    )
    try {
        $ex = $null
        $msg = $null
        if ($ErrorObject -and $ErrorObject.PSObject.Properties['Exception']) { $ex = $ErrorObject.Exception }
        if ($ErrorObject) { $msg = [string]$ErrorObject }
        if ([string]::IsNullOrWhiteSpace($msg) -and $ex) { $msg = [string]$ex.Message }

        if ($ex -and $ex.Response -and $ex.Response.StatusCode) {
            $code = [int]$ex.Response.StatusCode.value__
            if ($code -in 429,500,502,503,504) { return $true }
        }
        if (-not [string]::IsNullOrWhiteSpace($msg)) {
            if ($msg -match 'timed out|temporary|try again|forcibly closed|connection.*reset|could not be resolved|Something went wrong') { return $true }
        }
    } catch { }
    return $false
}

function Process-SiteForRetry {
    param(
        [Parameter(Mandatory = $true)] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [string] $TenantName,
        [Parameter(Mandatory = $true)] [string] $ClientId,
        [Parameter(Mandatory = $true)] [string] $CertificatePath,
        [Parameter(Mandatory = $false)] [securestring] $CertificatePassword,
        [Parameter(Mandatory = $true)] [string] $UserEmail,
        [Parameter(Mandatory = $true)] [object] $UserGroups,
        [Parameter(Mandatory = $false)] [bool] $IsQuiet = $false,
        [Parameter(Mandatory = $false)] [string] $LogPath
    )

    $results = @()
    try {
        if ($IsQuiet -and $LogPath) {
            [System.IO.File]::AppendAllText($LogPath, "$(Get-Date) INFO: Retrying $SiteUrl..." + [Environment]::NewLine)
        } else {
            Write-Host "$(Get-Date) INFO: Retrying $SiteUrl..."
        }

        if ($CertificatePassword) {
            $passwordBstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePassword)
            try { $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordBstr) } finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
            $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $plainPassword)
        } else {
            $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
        }
        $spoToken = Get-MsalToken -TenantId "$TenantName.onmicrosoft.com" -ClientId $ClientId -ClientCertificate $certificate -Scopes "https://$TenantName.sharepoint.com/.default"
        $headers = @{ Authorization = "Bearer $($spoToken.AccessToken)"; Accept = 'application/json;odata=nometadata' }

        # Site admins
        try {
            $adminsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$SiteUrl/_api/web/siteusers?`$filter=IsSiteAdmin%20eq%20true") -BaseHeaders $headers -Method GET
            $admins = @()
            if ($adminsResponse) { if ($adminsResponse.value) { $admins = $adminsResponse.value } else { $admins = @($adminsResponse) } }
            foreach ($admin in $admins) {
                $adminLogin = $admin.LoginName
                if ($adminLogin -match '\|') { $adminLogin = $adminLogin.Split('|')[-1] }
                if ($UserEmail -eq $adminLogin -or ($null -ne $admin.Email -and $UserEmail -eq $admin.Email) -or ($null -ne $UserGroups -and $UserGroups.GroupId -contains $adminLogin)) {
                    $results += [PSCustomObject]@{
                        UserPrincipalName = $UserEmail
                        SiteUrl           = $SiteUrl
                        SiteAdmin         = $true
                        GroupName         = $null
                        PermissionLevel   = $null
                        ListName          = $null
						ItemUrl           = $null
                        ListPermission    = $null
                        TotalRuntimeSeconds = $null
                    }
                    break
                }
            }
        } catch { }

        # Groups
        try {
            $groupsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$SiteUrl/_api/web/sitegroups") -BaseHeaders $headers -Method GET
            $siteGroups = @()
            if ($groupsResponse) { if ($groupsResponse.value) { $siteGroups = $groupsResponse.value } else { $siteGroups = @($groupsResponse) } }

            $webAssignmentsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$SiteUrl/_api/web/roleassignments?`$expand=Member,RoleDefinitionBindings") -BaseHeaders $headers -Method GET
            $webRoleAssignments = @()
            if ($webAssignmentsResponse) { if ($webAssignmentsResponse.value) { $webRoleAssignments = $webAssignmentsResponse.value } else { $webRoleAssignments = @($webAssignmentsResponse) } }

            foreach ($group in $siteGroups) {
                if ($null -eq $group) { continue }
                $groupTitle = [string]$group.Title
                if ($groupTitle -match "Limited Access|_catalog") { continue }

                if ($groupTitle -match '(?i)^\s*Sharing\s*Links') {
                    $matchingAssignments = @()
                    if ($webRoleAssignments) { $matchingAssignments = $webRoleAssignments | Where-Object { $_.Member -and ($_.Member.Title -eq $groupTitle) } }
                    $permSet = New-Object System.Collections.Generic.HashSet[string]
                    foreach ($ra in $matchingAssignments) { if ($ra.RoleDefinitionBindings) { foreach ($binding in $ra.RoleDefinitionBindings) { if ($binding.Name) { [void]$permSet.Add([string]$binding.Name) } } } }
                    $permString = (($permSet.ToArray()) -join ' | ')
                    if ([string]::IsNullOrWhiteSpace($permString)) { $permString = 'No Permissions' }
                    $results += [PSCustomObject]@{
                        UserPrincipalName = $UserEmail
                        SiteUrl           = $SiteUrl
                        SiteAdmin         = $false
                        GroupName         = $groupTitle
                        PermissionLevel   = $permString
                        ListName          = $null
						ItemUrl           = $null
                        ListPermission    = $null
                        TotalRuntimeSeconds = $null
                    }
                    continue
                }

                if (-not $group.Id -or [string]::IsNullOrWhiteSpace("$($group.Id)")) { continue }
                $membersResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$SiteUrl/_api/web/sitegroups/getbyid($($group.Id))/users") -BaseHeaders $headers -Method GET
                $groupMembers = @()
                if ($membersResponse) { if ($membersResponse.value) { $groupMembers = $membersResponse.value } else { $groupMembers = @($membersResponse) } }
                $userIsInGroup = $false
                foreach ($member in $groupMembers) {
                    if ($null -eq $member -or [string]::IsNullOrWhiteSpace([string]$member.LoginName)) { continue }
                    if ($member.LoginName -match '\\|') { $memberLogin = $member.LoginName.Split('|')[-1] } else { $memberLogin = $member.LoginName }
                    if ($UserEmail -eq $memberLogin -or ($null -ne $UserGroups -and $UserGroups.GroupId -contains $memberLogin)) { $userIsInGroup = $true; break }
                }
                if ($userIsInGroup) {
                    $bindingsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$SiteUrl/_api/web/roleassignments/getbyprincipalid($($group.Id))/roledefinitionbindings") -BaseHeaders $headers -Method GET
                    $permissions = @()
                    if ($bindingsResponse) { if ($bindingsResponse.value) { $permissions = $bindingsResponse.value } else { $permissions = @($bindingsResponse) } }
                    $permNames = @(); foreach ($p in $permissions) { $permNames += $p.Name }
                    $permString = ($permNames -join ' | ')
                    if ([string]::IsNullOrWhiteSpace($permString)) { $permString = 'No Permissions' }
                    $results += [PSCustomObject]@{
                        UserPrincipalName = $UserEmail
                        SiteUrl           = $SiteUrl
                        SiteAdmin         = $false
                        GroupName         = $groupTitle
                        PermissionLevel   = $permString
                        ListName          = $null
						ItemUrl           = $null
                        ListPermission    = $null
                        TotalRuntimeSeconds = $null
                    }
                }
            }
        } catch { }

        # Lists (document libraries only)
        try {
            $listsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$SiteUrl/_api/web/lists?`$select=Id,Title,BaseTemplate,Hidden&`$filter=BaseTemplate%20eq%20101%20and%20Hidden%20eq%20false&`$top=5000") -BaseHeaders $headers -Method GET
            $lists = @()
            if ($listsResponse) { if ($listsResponse.value) { $lists = $listsResponse.value } else { $lists = @($listsResponse) } }
            foreach ($list in $lists) {
                # List-level role assignments (only if list itself has unique permissions)
                try {
                    if ($list.HasUniqueRoleAssignments -eq $true) {
                $assignmentsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$SiteUrl/_api/web/lists(guid'$($list.Id)')/roleassignments?`$expand=Member,RoleDefinitionBindings") -BaseHeaders $headers -Method GET
                $assignments = @()
                if ($assignmentsResponse) { if ($assignmentsResponse.value) { $assignments = $assignmentsResponse.value } else { $assignments = @($assignmentsResponse) } }
                foreach ($ra in $assignments) {
                    $memberLogin = $ra.Member.LoginName
                    if ($memberLogin -match '\\|') { $memberLogin = $memberLogin.Split('|')[-1] }
                    if ($UserEmail -eq $memberLogin -or ($null -ne $UserGroups -and $UserGroups.GroupId -contains $memberLogin)) {
                        $permNames = @(); foreach ($b in $ra.RoleDefinitionBindings) { $permNames += $b.Name }
                        $permString = ($permNames -join ' | ')
                        $results += [PSCustomObject]@{
                            UserPrincipalName = $UserEmail
                            SiteUrl           = $SiteUrl
                            SiteAdmin         = $false
                            GroupName         = $null
                            PermissionLevel   = $null
                            ListName          = $list.Title
                                    ItemUrl           = $null
                            ListPermission    = $permString
                            TotalRuntimeSeconds = $null
                        }
                    }
                    if ($ra.Member -and $ra.Member.Title -like 'SharingLinks.*') {
                        $permNames = @(); foreach ($b in $ra.RoleDefinitionBindings) { $permNames += $b.Name }
                        $permString = ($permNames -join ' | ')
                        if ([string]::IsNullOrWhiteSpace($permString)) { $permString = 'No Permissions' }
                                $resolvedItemUrl = $null
                                if ($ResolveItemSharingLinks) {
                                    $resolvedItemUrl = Resolve-SharingLinkPrincipalToItemUrl -SiteUrl $SiteUrl -AccessToken $SpoAccessToken -PrincipalTitle $ra.Member.Title
                                }
                        $results += [PSCustomObject]@{
                            UserPrincipalName = $UserEmail
                            SiteUrl           = $SiteUrl
                            SiteAdmin         = $false
                            GroupName         = $ra.Member.Title
                            PermissionLevel   = $permString
                            ListName          = $list.Title
                                    ItemUrl           = $resolvedItemUrl
                            ListPermission    = $permString
                            TotalRuntimeSeconds = $null
                        }
                    }
                }
                    }
                } catch { }

                # Items with unique permissions (capture human-readable sharing link targets), REST-filtered and paged
                try { } catch { }
            }
        } catch { }
    } catch { }

    return $results
}

try {
Set-Location $PSScriptRoot

# Initialize optional logging file (no transcript to avoid lock when writing custom logs)
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

        # Create file if it doesn't exist
        if (-not (Test-Path -LiteralPath $Log)) { New-Item -ItemType File -Path $Log -Force | Out-Null }
    }
    catch {
        Write-Warn "Failed to prepare log file at '$Log': $($_.Exception.Message)"
    }
}

Write-Major "$(Get-Date) INFO: Starting processing for $UserEmail..." -ForegroundColor Green
Write-Detail "$(Get-Date) INFO: Connecting to tenant admin site..."
Write-Detail "$(Get-Date) INFO: Getting all site collections via Search REST..."
$siteCollections = Get-TenantSitesRest -UserEmail $UserEmail -ErrorAction Stop
Write-Detail "$(Get-Date) INFO: `tFound $($siteCollections.Count) site collections."

Write-Detail "$(Get-Date) INFO: Getting group membership for $UserEmail..."
$userGroupMembership = Get-UserGroupMembership -UserEmail $UserEmail -ErrorAction Stop
Write-Detail "$(Get-Date) INFO: `tFound $($userGroupMembership.Count) groups."

if (!$Append) {
    New-CsvFile -Path $CSVPath
}

# Create progress tracking
$totalSites = $siteCollections.Count

# Concurrency messaging (REST-based implementation supports parallel processing)
if ($ThrottleLimit -gt 1) {
    Write-Detail "$(Get-Date) INFO: Processing $totalSites sites with $ThrottleLimit parallel threads..."
} else {
    Write-Detail "$(Get-Date) INFO: Processing $totalSites sites sequentially..."
}

# Create synchronized hashtable for thread-safe operations
$syncHash = [hashtable]::Synchronized(@{
    ProcessedCount = 0
    TotalCount = $totalSites
    LogLines = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    RetrySites = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
})

# Acquire one SharePoint access token up-front and share with parallel runspaces
try {
    if ($CertificatePassword) {
        $passwordBstr_init = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePassword)
        try { $plainPassword_init = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordBstr_init) } finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr_init) }
        $certificate_init = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $plainPassword_init)
    } else {
        $certificate_init = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
    }
    $spoToken_init = Get-MsalToken -TenantId "$TenantName.onmicrosoft.com" -ClientId $ClientId -ClientCertificate $certificate_init -Scopes "https://$TenantName.sharepoint.com/.default"
    $SpoAccessToken = $spoToken_init.AccessToken
} catch {
    Write-Error "Failed to acquire initial SharePoint token: $($_.Exception.Message)"
    throw
}

# Process sites in parallel for improved performance
$parallelResults = $siteCollections | ForEach-Object -Parallel {
    # Use the pre-acquired SPO access token from parent runspace
    
    function Is-TransientSharePointError {
    	param(
    		[Parameter(Mandatory = $true)] [object] $ErrorObject
    	)
    	try {
    		$ex = $null
    		$msg = $null
    		if ($ErrorObject -and $ErrorObject.PSObject.Properties['Exception']) { $ex = $ErrorObject.Exception }
    		if ($ErrorObject) { $msg = [string]$ErrorObject }
    		if ([string]::IsNullOrWhiteSpace($msg) -and $ex) { $msg = [string]$ex.Message }

    		if ($ex -and $ex.Response -and $ex.Response.StatusCode) {
    			$code = [int]$ex.Response.StatusCode.value__
    			if ($code -in 429,500,502,503,504) { return $true }
    		}
    		if (-not [string]::IsNullOrWhiteSpace($msg)) {
    			if ($msg -match 'timed out|temporary|try again|forcibly closed|connection.*reset|could not be resolved|Something went wrong') { return $true }
    		}
    	} catch { }
    	return $false
    }
    
    function Invoke-SharePointRestWithAcceptFallback {
    	param (
    		[Parameter(Mandatory = $true)]
    		[string] $Uri,
    		[Parameter(Mandatory = $true)]
    		[hashtable] $BaseHeaders,
    		[Parameter(Mandatory = $false)]
    		[string] $Method = 'GET',
    		[Parameter(Mandatory = $false)]
    		[object] $Body
    	)

    	# Pinned handling with Search exception
    	$headers = @{}
    	foreach ($key in $BaseHeaders.Keys) { if ($key -ne 'Accept') { $headers[$key] = $BaseHeaders[$key] } }

    	$uriWithFormat = $Uri
    	$isSearch = $Uri -match '/_api/search/'
    	if ($isSearch) {
    		$accepts = @('application/json', 'application/json;odata=minimalmetadata', 'application/json;odata=verbose', '*/*', '')
    		$attempt = 0
    		while ($attempt -le $using:Max406Retries) {
    			$accept = $accepts[$attempt % $accepts.Count]
    			$variantHeaders = @{}
    			foreach ($k in $headers.Keys) { if ($k -ne 'Accept') { $variantHeaders[$k] = $headers[$k] } }
    			if ([string]::IsNullOrEmpty($accept)) { if ($variantHeaders.ContainsKey('Accept')) { $variantHeaders.Remove('Accept') } } else { $variantHeaders['Accept'] = $accept }
    			try {
    				if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
    					return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -Body $Body -ErrorAction Stop
    				} else {
    					return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -ErrorAction Stop
    				}
    			}
    			catch {
    				$resp = $_.Exception.Response
    				$status = $null
    				if ($resp -and $resp.StatusCode) { $status = [int]$resp.StatusCode.value__ }
    				$shouldNext = ($status -eq 406) -or ($_.Exception.Message -match '406|Not Acceptable')
    				if (-not $shouldNext) { throw }
    				$attempt++
    				continue
    			}
    		}
    		throw "Received 406 Not Acceptable from $Uri after $using:Max406Retries retries."
    	}

    	$uriBase = $Uri
    	$attempt = 0
    	while ($attempt -le $using:Max406Retries) {
    		# Variant 1: nometadata
    		$variantHeaders = @{}
    		foreach ($k in $headers.Keys) { if ($k -ne 'Accept') { $variantHeaders[$k] = $headers[$k] } }
    		$variantHeaders['Accept'] = 'application/json;odata=nometadata'
    		$uriWithFormat = $uriBase
    		if ($uriWithFormat -match '(\?|&)`?\$format=') {
    			$uriWithFormat = $uriWithFormat -replace '(\?|&)`?\$format=[^&]+', '$1`$format=application/json;odata=nometadata'
    		} else {
    			$separator = ($uriWithFormat -match '\?') ? '&' : '?'
    			$uriWithFormat = "$uriWithFormat${separator}`$format=application/json;odata=nometadata"
    		}
    		try {
    			if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
    				return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -Body $Body -ErrorAction Stop
    			} else {
    				return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -ErrorAction Stop
    			}
    		}
    		catch {
    			$resp = $_.Exception.Response
    			$status = $null
    			if ($resp -and $resp.StatusCode) { $status = [int]$resp.StatusCode.value__ }
    			$shouldNext = ($status -eq 406) -or ($_.Exception.Message -match '406|Not Acceptable')
    			if (-not $shouldNext) { throw }
    		}

    		# Variant 2: minimalmetadata
    		$variantHeaders = @{}
    		foreach ($k in $headers.Keys) { if ($k -ne 'Accept') { $variantHeaders[$k] = $headers[$k] } }
    		$variantHeaders['Accept'] = 'application/json;odata=minimalmetadata'
    		$uriWithFormat = $uriBase
    		if ($uriWithFormat -match '(\?|&)`?\$format=') {
    			$uriWithFormat = $uriWithFormat -replace '(\?|&)`?\$format=[^&]+', '$1`$format=application/json;odata=minimalmetadata'
    		} else {
    			$separator = ($uriWithFormat -match '\?') ? '&' : '?'
    			$uriWithFormat = "$uriWithFormat${separator}`$format=application/json;odata=minimalmetadata"
    		}
    		try {
    			if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
    				return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -Body $Body -ErrorAction Stop
    			} else {
    				return Invoke-RestMethod -Uri $uriWithFormat -Headers $variantHeaders -Method $Method -ErrorAction Stop
    			}
    		}
    		catch {
    			$resp = $_.Exception.Response
    			$status = $null
    			if ($resp -and $resp.StatusCode) { $status = [int]$resp.StatusCode.value__ }
    			$shouldNext = ($status -eq 406) -or ($_.Exception.Message -match '406|Not Acceptable')
    			if (-not $shouldNext) { throw }
    		}

    		$attempt++
    	}

    	throw "Received 406 Not Acceptable from $Uri after $using:Max406Retries retries."
    }

    $site = $_
    $localResults = @()
    
    # Import variables from parent scope (only those actually needed)
    $localUserEmail = $using:UserEmail
    $localUserGroups = $using:userGroupMembership
    $sync = $using:syncHash
    $localIsQuiet = $using:ConsoleQuiet
    $localLog = $using:Log
    
    # Increment and get the current count
    $sync.ProcessedCount++
    $currentIndex = $sync.ProcessedCount
    
    try {
        # Detailed per-site progress: log-only when -Log was supplied
        if ($localIsQuiet -and $localLog) {
            [void]$sync.LogLines.Add("$(Get-Date) INFO: Processing $($site.Url) ($currentIndex of $($sync.TotalCount))...")
        } else {
            Write-Host "$(Get-Date) INFO: Processing $($site.Url) ($currentIndex of $($sync.TotalCount))..."
        }
        
        # Check if user is site collection admin via SharePoint REST
        $isSiteAdmin = $false
        try {
            # Use shared SharePoint access token (app-only)
            $headers = @{ Authorization = "Bearer $using:SpoAccessToken"; Accept = 'application/json;odata=nometadata' }

            # Retrieve site admins
            $adminsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$($site.Url)/_api/web/siteusers?`$filter=IsSiteAdmin%20eq%20true") -BaseHeaders $headers -Method GET
            $admins = @()
            if ($adminsResponse) {
                if ($adminsResponse.value) { $admins = $adminsResponse.value } else { $admins = @($adminsResponse) }
            }

            foreach ($admin in $admins) {
                $adminLogin = $admin.LoginName
                if ($adminLogin -match '\|') { $adminLogin = $adminLogin.Split('|')[-1] }

                if ($localUserEmail -eq $adminLogin -or ($null -ne $admin.Email -and $localUserEmail -eq $admin.Email) -or ($null -ne $localUserGroups -and $localUserGroups.GroupId -contains $adminLogin)) {
                    $isSiteAdmin = $true
                    if ($localIsQuiet -and $localLog) {
                        [void]$sync.LogLines.Add("$(Get-Date) INFO: `t$localUserEmail is a site collection admin for $($site.Url).")
                    } else {
                        Write-Host "$(Get-Date) INFO: `t$localUserEmail is a site collection admin for $($site.Url)."
                    }
                    $localResults += [PSCustomObject]@{
                        UserPrincipalName = $localUserEmail
                        SiteUrl           = $site.Url
            SiteAdmin         = $true
            GroupName         = $null
            PermissionLevel   = $null
            ListName          = $null
            ItemUrl           = $null
            ListPermission    = $null
            TotalRuntimeSeconds = $null
        }
                    break
                }
            }
        }
        catch {
            if ($localIsQuiet -and $localLog) { [void]$sync.LogLines.Add("WARNING: Error checking site admin status for $($site.Url): $_") } else { Write-Warning "Error checking site admin status for $($site.Url): $_" }
            if (Is-TransientSharePointError -ErrorObject $_) { [void]$sync.RetrySites.Add(@{ Url = $site.Url }) }
        }
        
        if (-not $isSiteAdmin) {
            # Check SharePoint groups
            try {
                # Retrieve site groups via SharePoint REST
                $headersGroups = @{ Authorization = "Bearer $using:SpoAccessToken"; Accept = 'application/json;odata=nometadata' }
                $groupsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$($site.Url)/_api/web/sitegroups") -BaseHeaders $headersGroups -Method GET
                $siteGroups = @()
                if ($groupsResponse) { if ($groupsResponse.value) { $siteGroups = $groupsResponse.value } else { $siteGroups = @($groupsResponse) } }

                # Retrieve web-level role assignments to resolve SharingLinks principals
                $webAssignmentsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$($site.Url)/_api/web/roleassignments?`$expand=Member,RoleDefinitionBindings") -BaseHeaders $headersGroups -Method GET
                $webRoleAssignments = @()
                if ($webAssignmentsResponse) { if ($webAssignmentsResponse.value) { $webRoleAssignments = $webAssignmentsResponse.value } else { $webRoleAssignments = @($webAssignmentsResponse) } }
                foreach ($group in $siteGroups) {
                    try {
                        if ($null -eq $group) { continue }
                        $groupTitle = [string]$group.Title
                        # Skip system groups that might cause issues
                        if ($groupTitle -match "Limited Access|_catalog") {
        continue
    }

                        # Special handling for SharingLinks principals at web scope
                        if ($groupTitle -match '(?i)^\s*Sharing\s*Links') {
                            try {
                                $matchingAssignments = @()
                                if ($webRoleAssignments) { $matchingAssignments = $webRoleAssignments | Where-Object { $_.Member -and ($_.Member.Title -eq $groupTitle) } }
                                $permSet = New-Object System.Collections.Generic.HashSet[string]
                                foreach ($ra in $matchingAssignments) {
                                    if ($ra.RoleDefinitionBindings) { foreach ($binding in $ra.RoleDefinitionBindings) { if ($binding.Name) { [void]$permSet.Add([string]$binding.Name) } } }
                                }
                                $permString = (($permSet.ToArray()) -join ' | ')
                                if ([string]::IsNullOrWhiteSpace($permString)) { $permString = 'No Permissions' }
                                $resolvedItemUrl = $null

                                if ($localIsQuiet -and $localLog) {
                                    [void]$sync.LogLines.Add("$(Get-Date) INFO: `t$localUserEmail sharing link $groupTitle has $permString at site level.")
                                } else {
                                    Write-Host "$(Get-Date) INFO: `t$localUserEmail sharing link $groupTitle has $permString at site level."
                                }
                                $localResults += [PSCustomObject]@{
                                    UserPrincipalName = $localUserEmail
                                    SiteUrl           = $site.Url
                SiteAdmin         = $false
                                    GroupName         = $groupTitle
                                    PermissionLevel   = $permString
                ListName          = $null
                ItemUrl           = $resolvedItemUrl
                ListPermission    = $null
                TotalRuntimeSeconds = $null
            }
                                continue
                            } catch { continue }
                        }

                        # Ensure group has a valid Id before ID-based REST calls
                        if (-not $group.Id -or [string]::IsNullOrWhiteSpace("$($group.Id)")) { continue }

                        # Retrieve members via SharePoint REST
                        $membersResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$($site.Url)/_api/web/sitegroups/getbyid($($group.Id))/users") -BaseHeaders $headersGroups -Method GET
                        $groupMembers = @()
                        if ($membersResponse) { if ($membersResponse.value) { $groupMembers = $membersResponse.value } else { $groupMembers = @($membersResponse) } }
                        $userIsInGroup = $false
                        
                        foreach ($member in $groupMembers) {
                            if ($null -eq $member -or [string]::IsNullOrWhiteSpace([string]$member.LoginName)) { continue }
                            if ($member.LoginName -match '\\|') {
                                $memberLogin = $member.LoginName.Split('|')[-1]
                            } else {
                                $memberLogin = $member.LoginName
                            }
                            
                            if ($localUserEmail -eq $memberLogin -or ($null -ne $localUserGroups -and $localUserGroups.GroupId -contains $memberLogin)) {
                                $userIsInGroup = $true
                                break
                            }
                        }
                        
                        if ($userIsInGroup) {
                            # Retrieve group permissions via SharePoint REST
                            $bindingsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$($site.Url)/_api/web/roleassignments/getbyprincipalid($($group.Id))/roledefinitionbindings") -BaseHeaders $headersGroups -Method GET
                            $permissions = @()
                            if ($bindingsResponse) { if ($bindingsResponse.value) { $permissions = $bindingsResponse.value } else { $permissions = @($bindingsResponse) } }
                            $permString = ""
                            foreach ($perm in $permissions) {
                                $permString += $perm.Name + " | "
                            }
                            
                            if ($permString -eq "") {
                                $permString = "No Permissions"
                            } else {
                                $permString = $permString.Substring(0, $permString.Length - 3)
                            }
                            
                            if ($localIsQuiet -and $localLog) {
                                [void]$sync.LogLines.Add("$(Get-Date) INFO: `t$localUserEmail is a member of $groupTitle with $permString permissions.")
                            } else {
                                Write-Host "$(Get-Date) INFO: `t$localUserEmail is a member of $groupTitle with $permString permissions."
                            }
                            $localResults += [PSCustomObject]@{
                                UserPrincipalName = $localUserEmail
                                SiteUrl           = $site.Url
                SiteAdmin         = $false
                                GroupName         = $groupTitle
                                PermissionLevel   = $permString
                ListName          = $null
                ItemUrl           = $null
                ListPermission    = $null
                TotalRuntimeSeconds = $null
            }
                        }
                    }
                    catch {
                        if ($localIsQuiet -and $localLog) {
                            [void]$sync.LogLines.Add("WARNING: Error processing group $($groupTitle): $_")
                        } else {
                            Write-Warning "Error processing group $($groupTitle): $_"
                        }
                    }
                }
            }
            catch {
                if ($localIsQuiet -and $localLog) { [void]$sync.LogLines.Add("WARNING: Error getting site groups for $($site.Url): $_") } else { Write-Warning "Error getting site groups for $($site.Url): $_" }
                if (Is-TransientSharePointError -ErrorObject $_) { [void]$sync.RetrySites.Add(@{ Url = $site.Url }) }
            }
            
            # Check list permissions
            try {
                # Headers for REST calls
                $headersLists = @{ Authorization = "Bearer $using:SpoAccessToken"; Accept = 'application/json;odata=nometadata' }

                $excludedLists = @("App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", 
                    "Composed Looks", "Content and Structure Reports", "Content type publishing error log", 
                    "Converted Forms", "Device Channels", "Form Templates", "fpdatasources", 
                    "Get started with Apps for Office and SharePoint", "List Template Gallery", 
                    "Long Running Operation Status", "Maintenance Log Library", "Style Library", 
                    "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", 
                    "Quick Deploy Items", "Relationships List", "Reusable Content", 
                    "Search Config List", "Solution Gallery", "Site Collection Images", 
                    "Suggested Content Browser Locations", "TaxonomyHiddenList", 
                    "User Information List", "Web Part Gallery", "wfpub", "wfsvc", 
                    "Workflow History", "Workflow Tasks", "Preservation Hold Library", 
                    "SharePointHomeCacheList")

                # Document libraries only for performance
                $listsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$($site.Url)/_api/web/lists?`$select=Id,Title,HasUniqueRoleAssignments,BaseTemplate,Hidden&`$filter=BaseTemplate%20eq%20101%20and%20Hidden%20eq%20false&`$top=5000") -BaseHeaders $headersLists -Method GET
                $lists = @()
                if ($listsResponse) { if ($listsResponse.value) { $lists = $listsResponse.value } else { $lists = @($listsResponse) } }

                foreach ($list in $lists) {
                    try {
                        if ($excludedLists -contains $list.Title) { continue }

                        # Get role assignments for the list
                        $assignmentsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$($site.Url)/_api/web/lists(guid'$($list.Id)')/roleassignments?`$expand=Member,RoleDefinitionBindings") -BaseHeaders $headersLists -Method GET
                        $assignments = @()
                        if ($assignmentsResponse) { if ($assignmentsResponse.value) { $assignments = $assignmentsResponse.value } else { $assignments = @($assignmentsResponse) } }

                        foreach ($roleAssignment in $assignments) {
                            $memberLogin = $roleAssignment.Member.LoginName
                            if ($memberLogin -match '\\|') { $memberLogin = $memberLogin.Split('|')[-1] }

                            if ($localUserEmail -eq $memberLogin -or ($null -ne $localUserGroups -and $localUserGroups.GroupId -contains $memberLogin)) {
                                $permissionNames = @()
                                foreach ($binding in $roleAssignment.RoleDefinitionBindings) { $permissionNames += $binding.Name }
                                $permissionString = $permissionNames -join " | "

                                if ($localIsQuiet -and $localLog) {
                                    [void]$sync.LogLines.Add("$(Get-Date) INFO: `t$localUserEmail has $permissionString permissions on $($list.Title).")
                                } else {
                                    Write-Host "$(Get-Date) INFO: `t$localUserEmail has $permissionString permissions on $($list.Title)."
                                }
                                $localResults += [PSCustomObject]@{
                                    UserPrincipalName = $localUserEmail
                                    SiteUrl           = $site.Url
                SiteAdmin         = $false
                GroupName         = $null
                PermissionLevel   = $null
                                    ListName          = $list.Title
                                    ItemUrl           = $null
                                    ListPermission    = $permissionString
                TotalRuntimeSeconds = $null
            }
                            }

                            # Include SharingLinks principals at list scope
                            if ($roleAssignment.Member -and ($roleAssignment.Member.Title -match '(?i)^\s*Sharing\s*Links')) {
                                $permissionNames = @()
                                foreach ($binding in $roleAssignment.RoleDefinitionBindings) { $permissionNames += $binding.Name }
                                $permissionString = $permissionNames -join " | "
                                if ([string]::IsNullOrWhiteSpace($permissionString)) { $permissionString = 'No Permissions' }
                                $resolvedItemUrl = $null
                                if ($using:ResolveItemSharingLinks) {
                                    $resolvedItemUrl = Resolve-SharingLinkPrincipalToItemUrl -SiteUrl $site.Url -AccessToken $using:SpoAccessToken -PrincipalTitle $roleAssignment.Member.Title
                                    if (-not $resolvedItemUrl) {
                                        $resolvedItemUrl = Try-ResolveSharingLinkBySamplingItems -SiteUrl $site.Url -ListId $list.Id -PrincipalTitle $roleAssignment.Member.Title -MaxScan $using:MaxItemScanPerList -AccessToken $using:SpoAccessToken
                                    }
                                }

                                if ($localIsQuiet -and $localLog) {
                                    [void]$sync.LogLines.Add("$(Get-Date) INFO: `tSharing link $($roleAssignment.Member.Title) has $permissionString on list $($list.Title).")
                                } else {
                                    Write-Host "$(Get-Date) INFO: `tSharing link $($roleAssignment.Member.Title) has $permissionString on list $($list.Title)."
                                }
                                $localResults += [PSCustomObject]@{
                                    UserPrincipalName = $localUserEmail
                                    SiteUrl           = $site.Url
                SiteAdmin         = $false
                                    GroupName         = $roleAssignment.Member.Title
                                    PermissionLevel   = $permissionString
                ListName          = $list.Title
                ItemUrl           = $resolvedItemUrl
                ListPermission    = $permissionString
                TotalRuntimeSeconds = $null
            }
                            }
                        }
                        # Optional item-level enumeration to resolve item-level SharingLinks rows
                        if ($using:ResolveItemSharingLinks) {
                            try {
                                $page = 0
                                $next = "$($site.Url)/_api/web/lists(guid'$($list.Id)')/items?`$select=Id,FileRef,HasUniqueRoleAssignments&`$filter=HasUniqueRoleAssignments%20eq%20true&`$top=1000"
                                while ($next -and ($page*1000) -lt $using:MaxItemScanPerList) {
                                    $itemsResponse = Invoke-SharePointRestWithAcceptFallback -Uri $next -BaseHeaders $headersLists -Method GET
                                    $items = @()
                                    if ($itemsResponse) { if ($itemsResponse.value) { $items = $itemsResponse.value } else { $items = @($itemsResponse) } }
                                    foreach ($item in $items) {
                                        # For each item, fetch role assignments (cannot reliably expand via list item)
                                        $itemAssignmentsResponse = Invoke-SharePointRestWithAcceptFallback -Uri ("$($site.Url)/_api/web/lists(guid'$($list.Id)')/items($($item.Id))/roleassignments?`$expand=Member,RoleDefinitionBindings") -BaseHeaders $headersLists -Method GET
                                        $itemAssignments = @()
                                        if ($itemAssignmentsResponse) { if ($itemAssignmentsResponse.value) { $itemAssignments = $itemAssignmentsResponse.value } else { $itemAssignments = @($itemAssignmentsResponse) } }
                                        foreach ($ra in $itemAssignments) {
                                            if ($ra.Member -and ($ra.Member.Title -match '(?i)^\s*Sharing\s*Links')) {
                                                $permissionNames = @()
                                                if ($ra.RoleDefinitionBindings) { foreach ($b in $ra.RoleDefinitionBindings) { if ($b.Name) { $permissionNames += $b.Name } } }
                                                $permissionString = $permissionNames -join ' | '
                                                if ([string]::IsNullOrWhiteSpace($permissionString)) { $permissionString = 'No Permissions' }
                                                $su = [Uri]$site.Url
                                                $hostRoot = "$($su.Scheme)://$($su.Host)"
                                                $itemUrlAbsolute = $null
                                                if ($item.FileRef) { $itemUrlAbsolute = "$hostRoot$($item.FileRef)" }

                                                $localResults += [PSCustomObject]@{
                                                    UserPrincipalName = $localUserEmail
                                                    SiteUrl           = $site.Url
                SiteAdmin         = $false
                                                    GroupName         = $ra.Member.Title
                                                    PermissionLevel   = $permissionString
                ListName          = $list.Title
                ItemUrl           = $itemUrlAbsolute
                                                    ListPermission    = $null
                TotalRuntimeSeconds = $null
            }
                                            }
                                        }
                                        # Graph-based check for user-accessible sharing links on this item
                                        try {
                                            if ($item.FileRef) {
                                                $userLinks = Get-UserAccessibleSharingLinksForListItem -SiteUrl $site.Url -ListId $list.Id -ItemId $item.Id -UserEmail $localUserEmail -UserGroups $localUserGroups
                                                foreach ($ul in $userLinks) {
                                                    $localResults += [PSCustomObject]@{
                                                        UserPrincipalName = $localUserEmail
                                                        SiteUrl           = $site.Url
                            SiteAdmin         = $false
                                                        GroupName         = 'SharingLinks.Graph'
                                                        PermissionLevel   = $ul.Permission
                            ListName          = $list.Title
                            ItemUrl           = $ul.ItemUrl
                                                        ListPermission    = $null
                            TotalRuntimeSeconds = $null
                        }
                                                }
                                            }
                                        } catch { }
                                    }
                                    $page++
                                    $next = $null
                                    if ($itemsResponse -and $itemsResponse.PSObject.Properties['@odata.nextLink']) { $next = [string]$itemsResponse.'@odata.nextLink' }
                                }
                            } catch { }
                        }
                    }
                    catch {
                        # Silently skip lists that cause errors (could be system lists)
                        continue
                    }
                }
                # Last resort: search for items shared with the user and add those links
                if ($using:ResolveItemSharingLinks) {
                    try {
                        $sharedItems = Search-ItemsSharedWithUserRest -UserEmail $localUserEmail -FilterSiteUrl $site.Url
                        foreach ($si in $sharedItems) {
                            # Retrieve sharing information to populate permissions human-readably
                            $sharingInfo = Get-ItemSharingLinksRest -ItemUrl $si.ItemUrl -AccessToken $using:SpoAccessToken
                            $links = @()
                            if ($sharingInfo) {
                                if ($sharingInfo.links) { $links = $sharingInfo.links }
                                elseif ($sharingInfo.SharingLinks) { $links = $sharingInfo.SharingLinks }
                            }
                            if (-not $links -or $links.Count -eq 0) {
                                $localResults += [PSCustomObject]@{
                                    UserPrincipalName = $localUserEmail
                                    SiteUrl           = $site.Url
                SiteAdmin         = $false
                                    GroupName         = 'SharingLinks.SearchDerived'
                                    PermissionLevel   = $null
                ListName          = $null
                ItemUrl           = $si.ItemUrl
                                    ListPermission    = $null
                TotalRuntimeSeconds = $null
            }
                                continue
                            }
                            foreach ($link in $links) {
                                $roles = @()
                                if ($link.Roles) { foreach ($r in $link.Roles) { $roles += [string]$r } }
                                elseif ($link.Role) { $roles += [string]$link.Role }
                                $permString = ($roles -join ' | ')
                                if ([string]::IsNullOrWhiteSpace($permString)) { $permString = $link.LinkKind }
                                $localResults += [PSCustomObject]@{
                                    UserPrincipalName = $localUserEmail
                                    SiteUrl           = $site.Url
                SiteAdmin         = $false
                                    GroupName         = 'SharingLinks.SearchDerived'
                                    PermissionLevel   = $permString
                ListName          = $null
                ItemUrl           = $si.ItemUrl
                                    ListPermission    = $null
                TotalRuntimeSeconds = $null
            }
                            }
                        }
                    } catch { }
                }
            }
            catch {
                if ($localIsQuiet -and $localLog) { [void]$sync.LogLines.Add("WARNING: Error checking list permissions for $($site.Url): $_") } else { Write-Warning "Error checking list permissions for $($site.Url): $_" }
                if (Is-TransientSharePointError -ErrorObject $_) { [void]$sync.RetrySites.Add(@{ Url = $site.Url }) }
            }
        }
        
        # No PnP disconnect needed
    }
    catch {
        if ($localIsQuiet -and $localLog) { [void]$sync.LogLines.Add("WARNING: Error processing site $($site.Url): $($_.Exception.Message)") } else { Write-Warning "Error processing site $($site.Url): $($_.Exception.Message)" }
        if (Is-TransientSharePointError -ErrorObject $_) { [void]$sync.RetrySites.Add(@{ Url = $site.Url }) }
    }
    
    # Return results from this parallel iteration
    $localResults
    
} -ThrottleLimit $ThrottleLimit

# Final sweep for transient failures (single pass, no artificial waits)
$retrySitesUnique = @()
if ($syncHash.RetrySites.Count -gt 0) {
    $seen = New-Object System.Collections.Generic.HashSet[string]
    foreach ($it in $syncHash.RetrySites) {
        $u = [string]$it.Url
        if (-not [string]::IsNullOrWhiteSpace($u)) { if ($seen.Add($u)) { $retrySitesUnique += $u } }
    }
    if ($retrySitesUnique.Count -gt 0) {
        Write-Major "$(Get-Date) INFO: Retrying $($retrySitesUnique.Count) sites after main pass..."
        $retryResults = @()
        foreach ($retryUrl in $retrySitesUnique) {
            try {
                $retryResults += Process-SiteForRetry -SiteUrl $retryUrl -TenantName $TenantName -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -UserEmail $UserEmail -UserGroups $userGroupMembership -IsQuiet $ConsoleQuiet -LogPath $Log
            } catch { }
        }
        if ($retryResults -and $retryResults.Count -gt 0) {
            $parallelResults = @($parallelResults + $retryResults)
        }
    }
}

# Flush buffered detailed log lines to the log file (single write) if quiet mode is enabled
if ($ConsoleQuiet -and $Log -and ($syncHash.LogLines.Count -gt 0 -or $LogBuffer.Count -gt 0)) {
    try {
        # Write buffered detail lines from parallel runspaces in one operation
        $combined = New-Object System.Collections.Generic.List[string]
        if ($LogBuffer.Count -gt 0) { $combined.AddRange([string[]]$LogBuffer) }
        if ($syncHash.LogLines.Count -gt 0) { $combined.AddRange([string[]]$syncHash.LogLines) }
        if ($combined.Count -gt 0) {
            $all = ($combined.ToArray()) -join [Environment]::NewLine
            [System.IO.File]::AppendAllText($Log, $all + [Environment]::NewLine)
        }
        # reset local buffer for subsequent major messages
        if ($LogBuffer) { $LogBuffer.Clear() | Out-Null }
    } catch { }
}

# Deduplicate and write all results to CSV
Write-Major "$(Get-Date) INFO: Writing results to CSV..."
$dedup = @{}
foreach ($result in $parallelResults) {
    if ($result) {
        $key = ("$($result.UserPrincipalName)|$($result.SiteUrl)|$($result.SiteAdmin)|$($result.GroupName)|$($result.PermissionLevel)|$($result.ListName)|$($result.ItemUrl)|$($result.ListPermission)")
        if (-not $dedup.ContainsKey($key)) {
            $dedup[$key] = $true
            $result | Export-Csv -Path $CSVPath -Append -NoTypeInformation
        }
    }
}

# Append total runtime summary row for this user
$scriptEndTime = Get-Date
$elapsed = $scriptEndTime - $scriptStartTime
$totalSeconds = [math]::Round($elapsed.TotalSeconds, 2)
Write-Major "$(Get-Date) INFO: Total runtime for $($UserEmail): $($totalSeconds) seconds."

$csvLineObject = [PSCustomObject]@{
    UserPrincipalName   = $UserEmail
    SiteUrl             = $null
    SiteAdmin           = $null
    GroupName           = $null
    PermissionLevel     = $null
    ListName            = $null
    ItemUrl             = $null
    ListPermission      = $null
    TotalRuntimeSeconds = $totalSeconds
}
$csvLineObject | Export-Csv -Path $CSVPath -Append -NoTypeInformation

if ($transcriptStarted) {
    try { Stop-Transcript | Out-Null } catch { }
}

# Final flush for any remaining major messages
if ($ConsoleQuiet -and $Log -and $LogBuffer -and $LogBuffer.Count -gt 0) {
    try {
        $tail = ([string[]]$LogBuffer) -join [Environment]::NewLine
        [System.IO.File]::AppendAllText($Log, $tail + [Environment]::NewLine)
        $LogBuffer.Clear() | Out-Null
    } catch { }
}
}
catch {
    Write-Warn "Skipping user '$UserEmail' due to error: $($_.Exception.Message)"
    return
}

