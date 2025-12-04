# Get-SharePointFolderPermissions.ps1
# Description: This script enumerates all folders within a single SharePoint site's document libraries and exports them to a CSV file with permission details.
# Use case: Site content enumeration with permission reporting.

#requires -Version 7.5.4
param (
    [Parameter(Mandatory = $true)]
    [string] $SiteUrl,
    [Parameter(Mandatory = $true)]
    [string] $TenantName,
    [Parameter(Mandatory = $true)]
    [string] $ClientId,
    [Parameter(Mandatory = $true)]
    [string] $CertificatePath,
    [Parameter(Mandatory = $false)]
    [securestring] $CertificatePassword,
    [Parameter(Mandatory = $true)]
    [string] $CSVPath,
    [Parameter(Mandatory = $false)]
    [switch] $Append = $false,
    [Parameter(Mandatory = $false)]
    [string] $Log,
    [Parameter(Mandatory = $false)]
    [switch] $AppendLog,
    [Parameter(Mandatory = $false)]
    [int] $ThrottleLimit = 1,
    [Parameter(Mandatory = $false)]
    [int] $Max406Retries = 999
)

# Ensure MSAL.PS is available for the current user (non-interactive install on first run)
try {
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

# Console verbosity control: when -Log is supplied, keep console output minimal
$ConsoleQuiet = $false
if ($PSBoundParameters.ContainsKey('Log') -and $Log) { $ConsoleQuiet = $true }

# Buffer log lines when console is quiet
$LogBuffer = $null
if ($ConsoleQuiet -and $Log) {
    $LogBuffer = New-Object System.Collections.ArrayList
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)] [string] $Message,
        [Parameter(Mandatory = $true)] [string] $Level, # INFO, WARNING, ERROR, SUCCESS
        [Parameter(Mandatory = $false)] [System.ConsoleColor] $Color = 'Cyan'
    )
    
    $timestamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $logEntry = "$timestamp $Level`: $Message"

    # Determine console color based on Level if not explicitly overridden
    $consoleColor = $Color
    switch ($Level) {
        'WARNING' { $consoleColor = 'Yellow' }
        'ERROR'   { $consoleColor = 'Red' }
        'SUCCESS' { $consoleColor = 'Green' }
        'INFO'    { 
            if ($Color -eq 'Cyan') { $consoleColor = 'Cyan' } else { $consoleColor = $Color }
        }
    }

    if ($ConsoleQuiet -and $Log) {
        [void]$LogBuffer.Add($logEntry) 
    } else {
        Write-Host $logEntry -ForegroundColor $consoleColor 
    }
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

function New-CsvFile {
    <#
    .SYNOPSIS
    Creates a new CSV file with the folder schema.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $csv = [PSCustomObject]@{
        LibraryName             = $null
        FolderName              = $null
        "User/Group"            = $null
        Permission              = $null
        FolderUrl               = $null
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

function Get-DocumentLibraries {
    <#
    .SYNOPSIS
    Gets all document libraries (BaseTemplate 101) from the site.
    #>
    param (
        [Parameter(Mandatory = $true)] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [string] $AccessToken
    )

    $headers = @{ Authorization = "Bearer $AccessToken"; Accept = 'application/json;odata=nometadata' }
    $uri = "$SiteUrl/_api/web/lists?`$select=Id,Title,BaseTemplate,Hidden,RootFolder/ServerRelativeUrl&`$filter=BaseTemplate%20eq%20101%20and%20Hidden%20eq%20false&`$expand=RootFolder&`$top=5000"
    
    $response = Invoke-SharePointRestWithAcceptFallback -Uri $uri -BaseHeaders $headers -Method GET
    $libraries = @()
    if ($response) {
        if ($response.value) { $libraries = $response.value } else { $libraries = @($response) }
    }
    
    return $libraries
}

# Main script execution
try {
    Set-Location $PSScriptRoot

    # Initialize optional logging file
    if ($PSBoundParameters.ContainsKey('Log') -and $Log) {
        try {
            $logDirectory = Split-Path -Path $Log -Parent
            if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory)) {
                New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
            }

            if ((Test-Path -LiteralPath $Log) -and -not $AppendLog) {
                Remove-Item -LiteralPath $Log -Force -ErrorAction SilentlyContinue
            }

            if (-not (Test-Path -LiteralPath $Log)) { New-Item -ItemType File -Path $Log -Force | Out-Null }
        }
        catch {
            Write-Host "WARNING: Failed to prepare log file at '$Log': $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    Write-Log -Level "INFO" -Message "Starting folder enumeration for $SiteUrl..." -Color Green

    # Acquire SharePoint access token
    Write-Log -Level "INFO" -Message "Acquiring SharePoint access token..."
    $spoToken = Get-SharePointAccessToken
    $SpoAccessToken = $spoToken.AccessToken

    # Create CSV file
    if (!$Append) {
        New-CsvFile -Path $CSVPath
    }

    # Get document libraries
    Write-Log -Level "INFO" -Message "Discovering document libraries..."
    $libraries = Get-DocumentLibraries -SiteUrl $SiteUrl -AccessToken $SpoAccessToken
    Write-Log -Level "INFO" -Message "Found $($libraries.Count) document libraries." -Color Green

    # Create synchronized hashtable for thread-safe operations
    $syncHash = [hashtable]::Synchronized(@{
        ProcessedCount = 0
        TotalCount = $libraries.Count
        LogLines = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    })

    # Parse site URL for building full folder URLs
    $siteUri = [Uri]$SiteUrl
    $hostRoot = "$($siteUri.Scheme)://$($siteUri.Host)"

    # Process libraries in parallel
    Write-Log -Level "INFO" -Message "Processing libraries with ThrottleLimit=$ThrottleLimit..."

    $parallelResults = $libraries | ForEach-Object -Parallel {
        # Redefine helper functions inside parallel block
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

        $library = $_
        $sync = $using:syncHash
        $localSiteUrl = $using:SiteUrl
        $localAccessToken = $using:SpoAccessToken
        $localHostRoot = $using:hostRoot
        $localIsQuiet = $using:ConsoleQuiet
        $localLog = $using:Log

        $sync.ProcessedCount++
        $currentIndex = $sync.ProcessedCount

        # Re-implement minimal Write-Log for parallel context
        function Write-Log {
            param(
                [Parameter(Mandatory = $true)] [string] $Message,
                [Parameter(Mandatory = $true)] [string] $Level,
                [Parameter(Mandatory = $false)] [System.ConsoleColor] $Color = 'Cyan'
            )
            $timestamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
            $logEntry = "$timestamp $Level`: $Message"

            if ($localIsQuiet -and $localLog) {
                [void]$sync.LogLines.Add($logEntry)
            } else {
                $consoleColor = $Color
                switch ($Level) {
                    'WARNING' { $consoleColor = 'Yellow' }
                    'ERROR'   { $consoleColor = 'Red' }
                    'SUCCESS' { $consoleColor = 'Green' }
                    'INFO'    { if ($Color -eq 'Cyan') { $consoleColor = 'Cyan' } else { $consoleColor = $Color } }
                }
                Write-Host $logEntry -ForegroundColor $consoleColor
            }
        }

        try {
            Write-Log -Level "INFO" -Message "Processing library '$($library.Title)' ($currentIndex of $($sync.TotalCount))..."

            $headers = @{ Authorization = "Bearer $localAccessToken"; Accept = 'application/json;odata=nometadata' }

            # Track visited paths to prevent infinite recursion loops
            $visitedPaths = New-Object System.Collections.Generic.HashSet[string]
            
            # Helper function to recursively get all folders
            function Get-FoldersRecursively {
                param([string]$FolderServerRelativeUrl)
                
                # Safety checks to prevent infinite loops
                if ([string]::IsNullOrWhiteSpace($FolderServerRelativeUrl)) { return }
                if ($visitedPaths.Contains($FolderServerRelativeUrl)) { return }
                [void]$visitedPaths.Add($FolderServerRelativeUrl)

                # 1. Escape single quotes for OData (My'Folder -> My''Folder)
                $safePath = $FolderServerRelativeUrl -replace "'", "''"
                
                # 2. Encode for URL but PRESERVE SLASHES
                $encodedPath = [uri]::EscapeDataString($safePath) -replace '%2f', '/' -replace '%2F', '/'
                
                $foldersUrl = "$localSiteUrl/_api/web/GetFolderByServerRelativeUrl('$encodedPath')/Folders?`$select=Name,ServerRelativeUrl"
                
                try {
                    $response = Invoke-SharePointRestWithAcceptFallback -Uri $foldersUrl -BaseHeaders $headers -Method GET
                    $subFolders = @()
                    if ($response) {
                        if ($response.value) { $subFolders = $response.value } else { $subFolders = @($response) }
                    }
                    
                    foreach ($sf in $subFolders) {
                        # Skip system folders
                        if ([string]::IsNullOrWhiteSpace($sf.Name) -or $sf.Name -eq 'Forms' -or $sf.Name -like '_*') { continue }
                        
                        Write-Log -Level "INFO" -Message "`tFound folder: $($sf.ServerRelativeUrl)"
                        
                        # Fetch role assignments directly via ListItemAllFields endpoint to avoid needing the ID
                        try {
                            $folderSafePath = $sf.ServerRelativeUrl -replace "'", "''"
                            $folderEncodedPath = [uri]::EscapeDataString($folderSafePath) -replace '%2f', '/' -replace '%2F', '/'

                            $roleAssignmentsUrl = "$localSiteUrl/_api/web/GetFolderByServerRelativeUrl('$folderEncodedPath')/ListItemAllFields/RoleAssignments?`$expand=Member,RoleDefinitionBindings"
                            $raResponse = Invoke-SharePointRestWithAcceptFallback -Uri $roleAssignmentsUrl -BaseHeaders $headers -Method GET

                            $assignments = @()
                            if ($raResponse) {
                                if ($raResponse.value) { $assignments = $raResponse.value } else { $assignments = @($raResponse) }
                            }

                            # Group users by Permission Level
                            $permsGrouped = @{} # Key: PermissionName, Value: List of User/Groups

                            foreach ($ra in $assignments) {
                                $principalName = $null
                                if ($ra.Member) {
                                    # Prefer UPN for users, Title for groups
                                    if ($ra.Member.UserPrincipalName) {
                                        $principalName = $ra.Member.UserPrincipalName
                                    } elseif ($ra.Member.Title) {
                                        $principalName = $ra.Member.Title
                                    } elseif ($ra.Member.LoginName) {
                                        $principalName = $ra.Member.LoginName
                                    }
                                }
                                
                                if (-not $principalName) { continue }

                                if ($ra.RoleDefinitionBindings) {
                                    foreach ($binding in $ra.RoleDefinitionBindings) {
                                        $permName = $binding.Name
                                        if (-not $permsGrouped.ContainsKey($permName)) {
                                            $permsGrouped[$permName] = New-Object System.Collections.Generic.List[string]
                                        }
                                        [void]$permsGrouped[$permName].Add($principalName)
                                    }
                                }
                            }

                            $permissionsFound = $false
                            foreach ($permName in $permsGrouped.Keys) {
                                $usersList = ($permsGrouped[$permName] -join '; ')
                                $folderUrl = "$localHostRoot$($sf.ServerRelativeUrl)"
                                
                                Write-Output ([PSCustomObject]@{
                                    LibraryName             = $library.Title
                                    FolderName              = $sf.Name
                                    "User/Group"            = $usersList
                                    Permission              = $permName
                                    FolderUrl               = $folderUrl
                                })
                                $permissionsFound = $true
                            }

                            if (-not $permissionsFound) {
                                $folderUrl = "$localHostRoot$($sf.ServerRelativeUrl)"
                                Write-Output ([PSCustomObject]@{
                                    LibraryName             = $library.Title
                                    FolderName              = $sf.Name
                                    "User/Group"            = "No Access"
                                    Permission              = "None"
                                    FolderUrl               = $folderUrl
                                })
                            }
                            
                        } catch {
                            Write-Log -Level "WARNING" -Message "Failed to get permissions for '$($sf.ServerRelativeUrl)': $($_.Exception.Message)"
                            # Output basic info on error
                            $folderUrl = "$localHostRoot$($sf.ServerRelativeUrl)"
                            Write-Output ([PSCustomObject]@{
                                LibraryName             = $library.Title
                                FolderName              = $sf.Name
                                "User/Group"            = "Error"
                                Permission              = "Error"
                                FolderUrl               = $folderUrl
                            })
                        }
                        
                        # Recurse into subfolders
                        Get-FoldersRecursively -FolderServerRelativeUrl $sf.ServerRelativeUrl
                    }
                } catch {
                    Write-Log -Level "WARNING" -Message "Could not enumerate folder '$FolderServerRelativeUrl': $($_.Exception.Message)"
                }
            }
            
            # Get library root folder path and start recursion
            $libraryRootPath = $library.RootFolder.ServerRelativeUrl
            $localResults = @(Get-FoldersRecursively -FolderServerRelativeUrl $libraryRootPath)

            Write-Log -Level "INFO" -Message "`tFound $($localResults.Count) permission entries in '$($library.Title)'."
        }
        catch {
            Write-Log -Level "WARNING" -Message "Error processing library '$($library.Title)': $($_.Exception.Message)"
        }

        $localResults

    } -ThrottleLimit $ThrottleLimit

    # Flush buffered log lines
    if ($ConsoleQuiet -and $Log -and ($syncHash.LogLines.Count -gt 0 -or $LogBuffer.Count -gt 0)) {
        try {
            $combined = New-Object System.Collections.Generic.List[string]
            if ($LogBuffer.Count -gt 0) { $combined.AddRange([string[]]$LogBuffer) }
            if ($syncHash.LogLines.Count -gt 0) { $combined.AddRange([string[]]$syncHash.LogLines) }
            if ($combined.Count -gt 0) {
                $all = ($combined.ToArray()) -join [Environment]::NewLine
                [System.IO.File]::AppendAllText($Log, $all + [Environment]::NewLine)
            }
            if ($LogBuffer) { $LogBuffer.Clear() | Out-Null }
        } catch { }
    }

    # Write all results to CSV
    Write-Log -Level "INFO" -Message "Writing results to CSV..."
    $totalFolders = 0
    foreach ($result in $parallelResults) {
        if ($result) {
                $result | Export-Csv -Path $CSVPath -Append -NoTypeInformation
                $totalFolders++
        }
    }

    Write-Log -Level "SUCCESS" -Message "Found $totalFolders permission entries in total." -Color Green
    Write-Log -Level "SUCCESS" -Message "Folder enumeration complete. Results saved to $CSVPath" -Color Green
}
catch {
    Write-Log -Level "ERROR" -Message "Script failed: $($_.Exception.Message)" -Color Red
    throw
}
