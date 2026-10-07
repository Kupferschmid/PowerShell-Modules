# Invoke-ZEP
# Lokale Sitzungen lesen das Bearer-Token für die hhpberlin-ZEP-Instanz aus dem Windows Credential Manager.
# In Azure Automation Runbooks wird das Token aus Get-AutomationPSCredential gelesen.

# Version 0.3.1 07.10.2026 by Klaus Kupferschmid (tempero.it GmbH & hhpberlin GmbH)

#Requires -Modules @{ ModuleName = 'BetterCredentials'; ModuleVersion = '4.5' }

# Windows PowerShell 5.1 does not load System.Web (HttpUtility) by default.
Add-Type -AssemblyName System.Web

$script:ZEPSettings = [ordered]@{
    ServiceUserName = 'ZEP_hhpberlin'
    BaseUri = 'https://www.zep-online.de/zephhpberlin/next/api/v1'
    TokenTarget = 'ZEP_hhpberlin_BearerToken'
    CacheDirectory = (Join-Path $env:LOCALAPPDATA 'InvokeZEP')
    OfferCacheFileName = 'offers-cache.json'
    DefaultPageSize = 100
    DefaultThrottleDelaySeconds = 0.2
    DefaultCacheTtlMinutes = 60
    # ZEP returns local wall-clock times of this zone, although the values end with 'Z'.
    ServerTimeZone = 'Europe/Berlin'
}

function Update-ZEPDerivedSettings {
    $script:ZEP_cache_directory = $script:ZEPSettings.CacheDirectory
    $script:ZEP_offer_cache_path = Join-Path $script:ZEP_cache_directory $script:ZEPSettings.OfferCacheFileName
}

<#
.SYNOPSIS
Shows the active InvokeZEP module configuration.

.DESCRIPTION
Returns the current runtime settings for API endpoint, credential target,
cache location and default paging/throttle/cache values.

.EXAMPLE
Get-ZEPConfiguration

.EXAMPLE
Get-ZEPConfiguration | Format-List *

.EXAMPLE
Get-ZEPConfiguration | Select-Object BaseUri, DefaultPageSize, DefaultThrottleDelaySeconds, CacheDirectory
#>
function Get-ZEPConfiguration {
    [CmdletBinding()]
    param ()

    return [pscustomobject]@{
        ServiceUserName = $script:ZEPSettings.ServiceUserName
        BaseUri = $script:ZEPSettings.BaseUri
        TokenTarget = $script:ZEPSettings.TokenTarget
        CacheDirectory = $script:ZEPSettings.CacheDirectory
        OfferCacheFileName = $script:ZEPSettings.OfferCacheFileName
        DefaultPageSize = $script:ZEPSettings.DefaultPageSize
        DefaultThrottleDelaySeconds = $script:ZEPSettings.DefaultThrottleDelaySeconds
        DefaultCacheTtlMinutes = $script:ZEPSettings.DefaultCacheTtlMinutes
        ServerTimeZone = $script:ZEPSettings.ServerTimeZone
    }
}

<#
.SYNOPSIS
Updates InvokeZEP module configuration values.

.DESCRIPTION
Changes selected runtime settings for the current session. Only explicitly
provided parameters are updated.

.PARAMETER ServiceUserName
Credential name used in Azure Automation for Get-AutomationPSCredential.

.PARAMETER BaseUri
Base URI of the ZEP REST API.

.PARAMETER TokenTarget
Credential target name in Windows Credential Manager.

.PARAMETER CacheDirectory
Local directory for offer cache files.

.PARAMETER DefaultPageSize
Default page size for list operations.

.PARAMETER DefaultThrottleDelaySeconds
Default delay between REST requests.

.PARAMETER DefaultCacheTtlMinutes
Default cache TTL value for cache-related workflows.

.PARAMETER ServerTimeZone
Time zone of the wall-clock times returned by ZEP (IANA or Windows ID), default 'Europe/Berlin'.

.EXAMPLE
Set-ZEPConfiguration -DefaultPageSize 50

.EXAMPLE
Set-ZEPConfiguration -BaseUri 'https://www.zep-online.de/zephhpberlin/next/api/v1' -DefaultThrottleDelaySeconds 0.1

.EXAMPLE
Set-ZEPConfiguration -TokenTarget 'ZEP_hhpberlin_BearerToken' -ServiceUserName 'ZEP_hhpberlin'

.EXAMPLE
Set-ZEPConfiguration -CacheDirectory (Join-Path $env:LOCALAPPDATA 'InvokeZEP') -DefaultCacheTtlMinutes 120

.EXAMPLE
Set-ZEPConfiguration -DefaultPageSize 100 -DefaultThrottleDelaySeconds 0.2 -DefaultCacheTtlMinutes 60
#>
function Set-ZEPConfiguration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)][string]$ServiceUserName,
        [Parameter(Mandatory = $false)][string]$BaseUri,
        [Parameter(Mandatory = $false)][string]$TokenTarget,
        [Parameter(Mandatory = $false)][string]$CacheDirectory,
        [Parameter(Mandatory = $false)][int]$DefaultPageSize,
        [Parameter(Mandatory = $false)][double]$DefaultThrottleDelaySeconds,
        [Parameter(Mandatory = $false)][int]$DefaultCacheTtlMinutes,
        [Parameter(Mandatory = $false)][string]$ServerTimeZone
    )

    if ($PSBoundParameters.ContainsKey('ServiceUserName')) {
        $script:ZEPSettings.ServiceUserName = $ServiceUserName
    }

    if ($PSBoundParameters.ContainsKey('BaseUri')) {
        $script:ZEPSettings.BaseUri = $BaseUri.TrimEnd('/')
    }

    if ($PSBoundParameters.ContainsKey('TokenTarget')) {
        $script:ZEPSettings.TokenTarget = $TokenTarget
    }

    if ($PSBoundParameters.ContainsKey('CacheDirectory')) {
        $script:ZEPSettings.CacheDirectory = $CacheDirectory
    }

    if ($PSBoundParameters.ContainsKey('DefaultPageSize')) {
        $script:ZEPSettings.DefaultPageSize = $DefaultPageSize
    }

    if ($PSBoundParameters.ContainsKey('DefaultThrottleDelaySeconds')) {
        $script:ZEPSettings.DefaultThrottleDelaySeconds = $DefaultThrottleDelaySeconds
    }

    if ($PSBoundParameters.ContainsKey('DefaultCacheTtlMinutes')) {
        $script:ZEPSettings.DefaultCacheTtlMinutes = $DefaultCacheTtlMinutes
    }

    if ($PSBoundParameters.ContainsKey('ServerTimeZone')) {
        $null = Get-ZEPTimeZoneInfo -Id $ServerTimeZone
        $script:ZEPSettings.ServerTimeZone = $ServerTimeZone
        $script:ZEPTimeZoneInfo = $null
    }

    Update-ZEPDerivedSettings
    return Get-ZEPConfiguration
}

Update-ZEPDerivedSettings

function Initialize-AutomationEnvironment {
    if ($script:automationEnvironmentInitialized) {
        return
    }

    $script:automationEnvironmentInitialized = $true
    $script:env_runbook = $false

    try {
        if ($PSPrivateMetadata.JobId) {
            $script:env_runbook = $true
        }
    }
    catch {
        $script:env_runbook = $false
    }
}

function Convert-SecureStringToPlainText {
    param (
        [Parameter(Mandatory = $false)]
        [Security.SecureString] $SecureString
    )

    if ($null -eq $SecureString) {
        return ''
    }

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Get-StoredCredentialSafe {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Target
    )

    $globalErrorCountBefore = $global:Error.Count

    try {
        return Find-Credential -Filter $Target -ErrorAction Stop | Select-Object -First 1
    }
    catch {
        $message = $PSItem.Exception.Message
        $isExpectedNotFound = (
            $message -match 'Element.*(not found|nicht gefunden)' -or
            $message -match '(Credential|Element).*(not found|nicht gefunden)'
        )

        if ($isExpectedNotFound) {
            while ($global:Error.Count -gt $globalErrorCountBefore) {
                $global:Error.RemoveAt(0)
            }
            return $null
        }

        Write-Warning "Unerwarteter Fehler bei Find-Credential fuer Target '$Target': $message"
        return $null
    }
}

function Get-ZEPBearerToken {
    Initialize-AutomationEnvironment

    if ($env_runbook) {
        $automationCredentialCommand = Get-Command -Name Get-AutomationPSCredential -ErrorAction SilentlyContinue
        if (-not $automationCredentialCommand) {
            throw 'Get-AutomationPSCredential ist in dieser Runbook-Umgebung nicht verfuegbar.'
        }

        try {
            $credential = Get-AutomationPSCredential -Name $script:ZEPSettings.ServiceUserName -ErrorAction Stop
        }
        catch {
            throw "AutomationPSCredential mit dem Namen $($script:ZEPSettings.ServiceUserName) konnte nicht gelesen werden."
        }

        if (-not $credential) {
            throw "AutomationPSCredential mit dem Namen $($script:ZEPSettings.ServiceUserName) wurde nicht gefunden."
        }

        return Convert-SecureStringToPlainText -SecureString $credential.Password
    }

    $storedCredential = Get-StoredCredentialSafe -Target $script:ZEPSettings.TokenTarget
    if ($storedCredential) {
        return Convert-SecureStringToPlainText -SecureString $storedCredential.Password
    }

    Write-Host 'ZEP API-Token wird benoetigt' -ForegroundColor Yellow
    $credential = Microsoft.PowerShell.Security\Get-Credential -UserName $script:ZEPSettings.ServiceUserName -Message 'Geben Sie den ZEP API-Token ein'
    if (-not $credential) {
        throw 'ZEP API-Token wurde nicht eingegeben.'
    }

    Set-Credential -Target $script:ZEPSettings.TokenTarget -Credential $credential -Type Generic -Persistence Enterprise -Description 'ZEP API token' > $null
    return Convert-SecureStringToPlainText -SecureString $credential.Password
}

<#
.SYNOPSIS
Executes a ZEP REST API request.

.DESCRIPTION
Sends an HTTP request to the configured ZEP API endpoint with bearer token
authentication and retry handling for transient 429/503 responses.

.PARAMETER Path
Relative API path, for example 'offers' or 'offers/12345'.

.PARAMETER Method
HTTP method. Default is GET.

.PARAMETER Query
Hashtable of query parameters.

.PARAMETER Body
Request body object, serialized as JSON.

.PARAMETER RetryCount
Maximum retry attempts for transient errors.

.PARAMETER ThrottleDelaySeconds
Delay before each request execution.

.EXAMPLE
Invoke-ZEPRest -Path 'offers'

.EXAMPLE
Invoke-ZEPRest -Path 'offers' -Query @{ limit = 50; page = 1 }

.EXAMPLE
Invoke-ZEPRest -Path 'offers/12345' -Method 'GET' -RetryCount 3 -ThrottleDelaySeconds 0

.EXAMPLE
Invoke-ZEPRest -Path 'offers' -Method 'POST' -Body @{ name = 'Test' }
#>
function Invoke-ZEPRest {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][string]$Method = 'GET',
        [Parameter(Mandatory = $false)][hashtable]$Query,
        [Parameter(Mandatory = $false)]$Body,
        [Parameter(Mandatory = $false)][int]$RetryCount = 2,
        [Parameter(Mandatory = $false)][double]$ThrottleDelaySeconds = $script:ZEPSettings.DefaultThrottleDelaySeconds
    )

    $token = Get-ZEPBearerToken
    $headers = @{
        Accept = 'application/json'
        Authorization = "Bearer $token"
    }

    $resourceUri = $script:ZEPSettings.BaseUri.TrimEnd('/')
    if ($Path) {
        $resourceUri = "$resourceUri/$($Path.TrimStart('/'))"
    }

    $uriBuilder = [System.UriBuilder]::new($resourceUri)
    if ($Query -and $Query.Count -gt 0) {
        $queryCollection = [System.Web.HttpUtility]::ParseQueryString([string]::Empty)
        foreach ($key in $Query.Keys) {
            $value = $Query[$key]
            if ($null -eq $value) {
                continue
            }

            # Array values become repeated keys, e.g. id[]=1&id[]=2.
            if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
                foreach ($item in $value) {
                    $queryCollection.Add($key, [string]$item)
                }
                continue
            }

            $queryCollection[$key] = [string]$value
        }

        $queryString = $queryCollection.ToString()
        if (-not [string]::IsNullOrWhiteSpace($queryString)) {
            $uriBuilder.Query = $queryString
        }
    }

    $validatedUri = [uri]::new($uriBuilder.Uri.AbsoluteUri)

    $invokeParams = @{
        Uri = $validatedUri
        Method = $Method
        Headers = $headers
        ErrorAction = 'Stop'
    }

    if ($Body) {
        $invokeParams.ContentType = 'application/json'
        $invokeParams.Body = $Body | ConvertTo-Json -Depth 20
    }

    for ($attempt = 0; $attempt -le $RetryCount; $attempt++) {
        try {
            if ($ThrottleDelaySeconds -gt 0) {
                Start-Sleep -Milliseconds ([int]($ThrottleDelaySeconds * 1000))
            }

            return Invoke-RestMethod @invokeParams
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($attempt -lt $RetryCount -and ($statusCode -eq 429 -or $statusCode -eq 503)) {
                $backoffSeconds = [Math]::Min(10, [Math]::Pow(2, $attempt + 1))
                Start-Sleep -Seconds $backoffSeconds
                continue
            }

            throw
        }
    }
}

function ConvertTo-ZEPObject {
    param (
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )

    process {
        if ($null -eq $InputObject) {
            return $null
        }

        if ($InputObject -is [string] -or $InputObject -is [bool] -or $InputObject -is [int] -or $InputObject -is [long] -or $InputObject -is [double] -or $InputObject -is [decimal] -or $InputObject -is [datetime]) {
            return $InputObject
        }

        if ($InputObject -is [System.Collections.IDictionary]) {
            $result = [pscustomobject]::new()
            foreach ($key in $InputObject.Keys) {
                $value = $InputObject[$key]
                if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
                    $value = @($value | ForEach-Object { ConvertTo-ZEPObject -InputObject $_ })
                }
                elseif ($value -is [pscustomobject] -or $value -is [hashtable]) {
                    $value = ConvertTo-ZEPObject -InputObject $value
                }
                $result | Add-Member -NotePropertyName $key -NotePropertyValue $value
            }
            return $result
        }

        if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
            return @($InputObject | ForEach-Object { ConvertTo-ZEPObject -InputObject $_ })
        }

        $result = [pscustomobject]::new()
        foreach ($property in $InputObject.PSObject.Properties) {
            $value = $property.Value
            if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
                $value = @($value | ForEach-Object { ConvertTo-ZEPObject -InputObject $_ })
            }
            elseif ($value -is [pscustomobject] -or $value -is [hashtable]) {
                $value = ConvertTo-ZEPObject -InputObject $value
            }
            $result | Add-Member -NotePropertyName $property.Name -NotePropertyValue $value
        }

        return $result
    }
}

function Get-ZEPResponseItems {
    param (
        [Parameter(Mandatory = $false)]$Payload
    )

    if ($null -eq $Payload) {
        return @()
    }

    if ($Payload -is [System.Collections.IEnumerable] -and -not ($Payload -is [string]) -and -not ($Payload -is [System.Collections.IDictionary])) {
        return @($Payload)
    }

    if ($Payload.PSObject.Properties.Name -contains 'data') {
        return Get-ZEPResponseItems -Payload $Payload.data
    }

    if ($Payload.PSObject.Properties.Name -contains 'items') {
        return Get-ZEPResponseItems -Payload $Payload.items
    }

    if ($Payload.PSObject.Properties.Name -contains 'offers') {
        return Get-ZEPResponseItems -Payload $Payload.offers
    }

    return @($Payload)
}

function Get-ZEPResponseTotal {
    param (
        [Parameter(Mandatory = $false)]$Payload
    )

    if ($null -eq $Payload) {
        return $null
    }

    if ($Payload -is [System.Collections.IDictionary]) {
        if ($Payload.Contains('meta')) {
            $meta = $Payload['meta']
            if ($meta -is [System.Collections.IDictionary] -and $meta.Contains('total')) {
                return [int]$meta['total']
            }

            if ($meta -and $meta.PSObject.Properties.Name -contains 'total') {
                return [int]$meta.total
            }
        }

        if ($Payload.Contains('data')) {
            return Get-ZEPResponseTotal -Payload $Payload.data
        }

        if ($Payload.Contains('total')) {
            return [int]$Payload.total
        }

        return $null
    }

    if ($Payload.PSObject.Properties.Name -contains 'meta') {
        $meta = $Payload.meta
        if ($meta -and $meta.PSObject.Properties.Name -contains 'total') {
            return [int]$meta.total
        }
    }

    if ($Payload.PSObject.Properties.Name -contains 'data') {
        return Get-ZEPResponseTotal -Payload $Payload.data
    }

    if ($Payload.PSObject.Properties.Name -contains 'total') {
        return [int]$Payload.total
    }

    return $null
}

function Test-ZEPCacheAgeExceeded {
    param (
        [Parameter(Mandatory = $false)]$CacheObject,
        [Parameter(Mandatory = $false)][int]$CacheTtlMinutes = $script:ZEPSettings.DefaultCacheTtlMinutes
    )

    if (-not $CacheObject) {
        return $true
    }

    if ($CacheTtlMinutes -le 0) {
        return $true
    }

    if (-not ($CacheObject.PSObject.Properties.Name -contains 'generatedAtUtc')) {
        return $true
    }

    $rawTimestamp = $CacheObject.generatedAtUtc
    if ($null -eq $rawTimestamp) {
        return $true
    }

    $generatedAt = [datetimeoffset]::MinValue
    if ($rawTimestamp -is [datetimeoffset]) {
        $generatedAt = [datetimeoffset]$rawTimestamp
    }
    elseif ($rawTimestamp -is [datetime]) {
        $generatedAt = [datetimeoffset]([datetime]$rawTimestamp).ToUniversalTime()
    }
    else {
        $rawTimestampText = [string]$rawTimestamp
        if ([string]::IsNullOrWhiteSpace($rawTimestampText)) {
            return $true
        }

        if (-not [datetimeoffset]::TryParseExact(
                $rawTimestampText,
                'o',
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::RoundtripKind,
                [ref]$generatedAt
            )) {
            if (-not [datetimeoffset]::TryParse($rawTimestampText, [ref]$generatedAt)) {
                return $true
            }
        }
    }

    $ageMinutes = ((Get-Date).ToUniversalTime() - $generatedAt.UtcDateTime).TotalMinutes
    return $ageMinutes -ge $CacheTtlMinutes
}

<#
.SYNOPSIS
Checks whether the offers cache should be refreshed.

.DESCRIPTION
Validates the cache by comparing totals and item counts against either a given
CurrentTotal or a lightweight API validation request.

.PARAMETER CacheObject
Cache object loaded from the offer cache file.

.PARAMETER CurrentTotal
Optional externally provided current total.

.PARAMETER CacheTtlMinutes
Compatibility parameter for cache policy workflows.

.EXAMPLE
$cache = Read-ZEPOfferCache
Test-ZEPOffersCacheNeedsRefresh -CacheObject $cache

.EXAMPLE
$cache = Read-ZEPOfferCache
Test-ZEPOffersCacheNeedsRefresh -CacheObject $cache -CurrentTotal 250

.EXAMPLE
Test-ZEPOffersCacheNeedsRefresh -CacheObject $null
#>
function Test-ZEPOffersCacheNeedsRefresh {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]$CacheObject,
        [Parameter(Mandatory = $false)][int]$CurrentTotal,
        [Parameter(Mandatory = $false)][int]$CacheTtlMinutes = $script:ZEPSettings.DefaultCacheTtlMinutes
    )

    if (-not $CacheObject) {
        return $true
    }

    if (Test-ZEPCacheAgeExceeded -CacheObject $CacheObject -CacheTtlMinutes $CacheTtlMinutes) {
        return $true
    }

    if ($PSBoundParameters.ContainsKey('CurrentTotal')) {
        $totalChanged = [int]$CacheObject.total -ne [int]$CurrentTotal
        if ($totalChanged) {
            return $true
        }

        $cachedItems = @($CacheObject.items)
        $itemCount = @($cachedItems).Count
        return $itemCount -ne [int]$CurrentTotal
    }

    $cachedTotal = $null
    if ($CacheObject.PSObject.Properties.Name -contains 'total' -and $null -ne $CacheObject.total) {
        $cachedTotal = [int]$CacheObject.total
    }

    if ($null -eq $cachedTotal) {
        return $true
    }

    try {
        $validationResponse = Invoke-ZEPRest -Path 'offers' -Query @{ limit = 1; page = 1 } -ThrottleDelaySeconds 0
        $validationTotal = Get-ZEPResponseTotal -Payload $validationResponse
        if ($null -eq $validationTotal) {
            return $false
        }

        if ([int]$validationTotal -ne [int]$cachedTotal) {
            return $true
        }

        $cachedItems = @($CacheObject.items)
        return @($cachedItems).Count -ne [int]$validationTotal
    }
    catch {
        return $false
    }
}

function Get-ZEPDefaultOfferCachePath {
    return $script:ZEP_offer_cache_path
}

function Initialize-ZEPCacheStore {
    param (
        [Parameter(Mandatory = $false)][string]$Path = $script:ZEP_offer_cache_path
    )

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
}

function ConvertTo-ZEPFingerprint {
    param (
        [Parameter(Mandatory = $true)]$Record
    )

    $canonicalJson = $Record | ConvertTo-Json -Depth 20 -Compress
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($canonicalJson)
    $hash = $sha256.ComputeHash($bytes)
    return [System.BitConverter]::ToString($hash).Replace('-', '').ToLowerInvariant()
}

function Read-ZEPOfferCache {
    param (
        [Parameter(Mandatory = $false)][string]$Path = $script:ZEP_offer_cache_path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        $json = Get-Content -LiteralPath $Path -Raw
        # ConvertFrom-Json -Depth exists only from PowerShell 6.2 on.
        if ($PSVersionTable.PSVersion -ge [version]'6.2') {
            return $json | ConvertFrom-Json -Depth 100
        }

        return $json | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Write-ZEPOfferCache {
    param (
        [Parameter(Mandatory = $true)]$CacheObject,
        [Parameter(Mandatory = $false)][string]$Path = $script:ZEP_offer_cache_path
    )

    Initialize-ZEPCacheStore -Path $Path
    $CacheObject | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $Path -Encoding UTF8
}

<#
.SYNOPSIS
Gets a single ZEP offer by ID.

.DESCRIPTION
Returns one offer from cache (when available) or from the API, and writes the
normalized result back to cache.

.PARAMETER Id
Offer ID.

.PARAMETER UseCache
Enables cache read/write behavior.

.PARAMETER Refresh
Forces API retrieval and bypasses cache read.

.PARAMETER CacheTtlMinutes
Compatibility parameter for cache policy workflows.

.PARAMETER CachePath
Path to offer cache JSON file.

.EXAMPLE
Get-ZEPOffer -Id 12345

.EXAMPLE
Get-ZEPOffer -Id 12345 -UseCache

.EXAMPLE
Get-ZEPOffer -Id 12345 -Refresh

.EXAMPLE
Get-ZEPOffer -Id 12345 -UseCache -CachePath (Join-Path $env:LOCALAPPDATA 'InvokeZEP\\offers-cache.json')

.EXAMPLE
Get-ZEPOffer -Id 12345 -UseCache -CacheTtlMinutes 60
#>
function Get-ZEPOffer {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][int]$Id,
        [Parameter(Mandatory = $false)][switch]$UseCache,
        [Parameter(Mandatory = $false)][switch]$Refresh,
        [Parameter(Mandatory = $false)][int]$CacheTtlMinutes = $script:ZEPSettings.DefaultCacheTtlMinutes,
        [Parameter(Mandatory = $false)][string]$CachePath = $script:ZEP_offer_cache_path
    )

    $shouldUseCache = $true
    if ($PSBoundParameters.ContainsKey('UseCache')) {
        $shouldUseCache = $UseCache.IsPresent
    }
    if ($Refresh) {
        $shouldUseCache = $false
    }

    $cache = $null
    if ($shouldUseCache) {
        $cache = Read-ZEPOfferCache -Path $CachePath
        if ($cache -and $cache.items -and -not (Test-ZEPCacheAgeExceeded -CacheObject $cache -CacheTtlMinutes $CacheTtlMinutes)) {
            $cachedEntry = $cache.items | Where-Object { [int]$_.id -eq $Id } | Select-Object -First 1
            if ($cachedEntry -and $cachedEntry.data) {
                return $cachedEntry.data
            }
        }
    }

    $response = Invoke-ZEPRest -Path "offers/$Id"
    $items = Get-ZEPResponseItems -Payload $response
    $record = $items | Select-Object -First 1

    if ($record) {
        $normalizedRecord = ConvertTo-ZEPObject -InputObject $record
        if ($shouldUseCache) {
            $cache = Read-ZEPOfferCache -Path $CachePath
            if (-not $cache) {
                $cache = [pscustomobject]@{ version = 1; generatedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); total = $null; items = @() }
            }

            $existingEntries = @($cache.items)
            $updatedEntries = @()
            $updated = $false
            foreach ($entry in $existingEntries) {
                if ([int]$entry.id -eq $Id) {
                    $updatedEntries += [pscustomobject]@{
                        id = $Id
                        data = $normalizedRecord
                        fingerprint = ConvertTo-ZEPFingerprint -Record $normalizedRecord
                    }
                    $updated = $true
                }
                else {
                    $updatedEntries += $entry
                }
            }

            if (-not $updated) {
                $updatedEntries += [pscustomobject]@{
                    id = $Id
                    data = $normalizedRecord
                    fingerprint = ConvertTo-ZEPFingerprint -Record $normalizedRecord
                }
            }

            $cache.items = @($updatedEntries)
            $cache.generatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
            Write-ZEPOfferCache -CacheObject $cache -Path $CachePath
        }

        return $normalizedRecord
    }

    return $null
}

<#
.SYNOPSIS
Gets all ZEP offers.

.DESCRIPTION
Loads all offers from cache when valid, or fetches paged results from the API.
Optionally validates cache freshness before reuse.

.PARAMETER UseCache
Enables cache read/write behavior.

.PARAMETER Refresh
Forces API retrieval and bypasses cache read.

.PARAMETER ValidateCache
Validates cache against API total count before reuse.

.PARAMETER PageSize
Page size for API pagination.

.PARAMETER CacheTtlMinutes
Compatibility parameter for cache policy workflows.

.PARAMETER CachePath
Path to offer cache JSON file.

.EXAMPLE
Get-ZEPOffers

.EXAMPLE
Get-ZEPOffers -UseCache

.EXAMPLE
Get-ZEPOffers -UseCache -ValidateCache

.EXAMPLE
Get-ZEPOffers -Refresh -PageSize 200

.EXAMPLE
Get-ZEPOffers -UseCache -CachePath (Join-Path $env:LOCALAPPDATA 'InvokeZEP\\offers-cache.json')

.EXAMPLE
Get-ZEPOffers -UseCache -ValidateCache -CacheTtlMinutes 60
#>
function Get-ZEPOffers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)][switch]$UseCache,
        [Parameter(Mandatory = $false)][switch]$Refresh,
        [Parameter(Mandatory = $false)][switch]$ValidateCache,
        [Parameter(Mandatory = $false)][int]$PageSize = $script:ZEPSettings.DefaultPageSize,
        [Parameter(Mandatory = $false)][int]$CacheTtlMinutes = $script:ZEPSettings.DefaultCacheTtlMinutes,
        [Parameter(Mandatory = $false)][string]$CachePath = $script:ZEP_offer_cache_path
    )

    $shouldUseCache = $true
    if ($PSBoundParameters.ContainsKey('UseCache')) {
        $shouldUseCache = $UseCache.IsPresent
    }
    if ($Refresh) {
        $shouldUseCache = $false
    }

    $shouldValidateCache = $true
    if ($PSBoundParameters.ContainsKey('ValidateCache')) {
        $shouldValidateCache = $ValidateCache.IsPresent
    }

    $cacheInvalidReason = $null
    if ($shouldUseCache) {
        $cache = Read-ZEPOfferCache -Path $CachePath
        if ($cache -and $cache.items) {
            if (Test-ZEPCacheAgeExceeded -CacheObject $cache -CacheTtlMinutes $CacheTtlMinutes) {
                $cacheInvalidReason = "Cache ungültig: älter als $CacheTtlMinutes Minuten."
                Write-Host $cacheInvalidReason -ForegroundColor Yellow
            }

            $cachedTotal = $null
            if ($cache.PSObject.Properties.Name -contains 'total' -and $null -ne $cache.total) {
                $cachedTotal = [int]$cache.total
            }

            $needsRefresh = $false
            if ($cacheInvalidReason) {
                $needsRefresh = $true
            }
            if ($shouldValidateCache) {
                try {
                    $validationResponse = Invoke-ZEPRest -Path 'offers' -Query @{ limit = 1; page = 1 } -ThrottleDelaySeconds 0
                    $validationTotal = Get-ZEPResponseTotal -Payload $validationResponse
                    if ($null -ne $validationTotal -and $null -ne $cachedTotal -and [int]$validationTotal -ne [int]$cachedTotal) {
                        $needsRefresh = $true
                        $cacheInvalidReason = "Cache ungültig: Gesamtanzahl der Angebote hat sich geändert (cached=$cachedTotal, aktuell=$validationTotal)."
                    }
                }
                catch {
                    $needsRefresh = $false
                }
            }

            if (-not $needsRefresh) {
                return @($cache.items | ForEach-Object { $_.data })
            }

            if (-not $cacheInvalidReason) {
                $cacheInvalidReason = 'Cache wird nicht verwendet, weil eine Validierung oder Aktualisierung erforderlich ist.'
            }
            Write-Host $cacheInvalidReason -ForegroundColor Yellow
        }
        else {
            $cacheInvalidReason = 'Cache ist leer oder enthält keine gültigen Einträge.'
            Write-Host $cacheInvalidReason -ForegroundColor Yellow
        }
    }

    $allRecords = [System.Collections.Generic.List[object]]::new()
    $page = 1
    $collectionTotal = $null
    while ($true) {
        $query = @{ limit = $PageSize; page = $page }

        $response = Invoke-ZEPRest -Path 'offers' -Query $query
        if ($null -eq $collectionTotal) {
            $collectionTotal = Get-ZEPResponseTotal -Payload $response
        }

        $pageItems = Get-ZEPResponseItems -Payload $response
        if (-not $pageItems -or $pageItems.Count -eq 0) {
            break
        }

        foreach ($pageItem in $pageItems) {
            $allRecords.Add((ConvertTo-ZEPObject -InputObject $pageItem))
        }

        if ($null -ne $collectionTotal) {
            Write-Host ("Fetched page $page with $($pageItems.Count) items. Total fetched so far: $($allRecords.Count) of $collectionTotal.")
        }
        else {
            Write-Host ("Fetched page $page with $($pageItems.Count) items. Total fetched so far: $($allRecords.Count).")
        }

        $lastPage = $null
        if ($response.PSObject.Properties.Name -contains 'meta') {
            $meta = $response.meta
            if ($meta -and $meta.PSObject.Properties.Name -contains 'last_page') {
                $lastPage = [int]$meta.last_page
            }
        }

        if ($null -ne $lastPage -and $page -ge $lastPage) {
            break
        }

        if ($pageItems.Count -lt $PageSize -and $null -eq $lastPage) {
            break
        }
        $page++
    }

    $cacheObject = [pscustomobject]@{
        version = 1
        generatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        total = if ($null -ne $collectionTotal) { [int]$collectionTotal } else { $allRecords.Count }
        items = @()
    }

    foreach ($record in $allRecords) {
        $cacheObject.items += [pscustomobject]@{
            id = $record.id
            data = $record
            fingerprint = ConvertTo-ZEPFingerprint -Record $record
        }
    }

    Write-ZEPOfferCache -CacheObject $cacheObject -Path $CachePath
    return @($allRecords)
}

<#
.SYNOPSIS
Gets offer items for a specific ZEP offer.

.DESCRIPTION
Returns one page of offer items directly from the API.

.PARAMETER OfferId
Offer ID for which items are requested.

.PARAMETER PageSize
Page size for item paging.

.PARAMETER Page
Page number to retrieve.

.EXAMPLE
Get-ZEPOfferItems -OfferId 12345

.EXAMPLE
Get-ZEPOfferItems -OfferId 12345 -Page 2

.EXAMPLE
Get-ZEPOfferItems -OfferId 12345 -PageSize 200 -Page 1
#>
function Get-ZEPOfferItems {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][int]$OfferId,
        [Parameter(Mandatory = $false)][int]$PageSize = $script:ZEPSettings.DefaultPageSize,
        [Parameter(Mandatory = $false)][int]$Page = 1
    )

    $response = Invoke-ZEPRest -Path "offers/$OfferId/items" -Query @{ limit = $PageSize; page = $Page }
    $pageItems = Get-ZEPResponseItems -Payload $response
    $result = @($pageItems | ForEach-Object { ConvertTo-ZEPObject -InputObject $_ })
    return $result
}

function Get-ZEPPropertyPathsFromObject {
    param (
        [Parameter(Mandatory = $false)]$InputObject,
        [Parameter(Mandatory = $false)][string]$Prefix = '',
        [Parameter(Mandatory = $false)][System.Collections.Generic.HashSet[string]]$PropertySet,
        [Parameter(Mandatory = $false)][int]$MaxDepth = 10
    )

    if ($null -eq $PropertySet) {
        return
    }

    if ($null -eq $InputObject) {
        return
    }

    $isScalar = {
        param($Value)
        if ($null -eq $Value) { return $true }
        if (
            $Value -is [string] -or
            $Value -is [bool] -or
            $Value -is [int] -or
            $Value -is [long] -or
            $Value -is [double] -or
            $Value -is [decimal] -or
            $Value -is [datetime] -or
            $Value -is [datetimeoffset] -or
            $Value -is [timespan] -or
            $Value -is [guid]
        ) {
            return $true
        }

        $valueType = $Value.GetType()
        return $valueType.IsEnum
    }

    $stack = [System.Collections.Generic.List[object]]::new()
    $stack.Add([pscustomobject]@{ Value = $InputObject; Prefix = $Prefix; Depth = 0 })

    while ($stack.Count -gt 0) {
        $node = $stack[$stack.Count - 1]
        $stack.RemoveAt($stack.Count - 1)

        $value = $node.Value
        $prefixValue = [string]$node.Prefix
        $depth = [int]$node.Depth

        if ($null -eq $value -or (& $isScalar -Value $value)) {
            continue
        }

        if ($depth -ge $MaxDepth) {
            continue
        }

        if ($value -is [System.Collections.IDictionary]) {
            foreach ($key in $value.Keys) {
                $name = if ([string]::IsNullOrWhiteSpace($prefixValue)) { [string]$key } else { "$prefixValue.$key" }
                [void]$PropertySet.Add($name)

                $childValue = $value[$key]
                if ($null -ne $childValue -and -not (& $isScalar -Value $childValue)) {
                    $stack.Add([pscustomobject]@{ Value = $childValue; Prefix = $name; Depth = ($depth + 1) })
                }
            }
            continue
        }

        if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
            foreach ($entry in $value) {
                if ($null -ne $entry -and -not (& $isScalar -Value $entry)) {
                    $stack.Add([pscustomobject]@{ Value = $entry; Prefix = $prefixValue; Depth = ($depth + 1) })
                }
            }
            continue
        }

        foreach ($property in $value.PSObject.Properties) {
            $name = if ([string]::IsNullOrWhiteSpace($prefixValue)) { [string]$property.Name } else { "$prefixValue.$($property.Name)" }
            [void]$PropertySet.Add($name)

            $childValue = $property.Value
            if ($null -ne $childValue -and -not (& $isScalar -Value $childValue)) {
                $stack.Add([pscustomobject]@{ Value = $childValue; Prefix = $name; Depth = ($depth + 1) })
            }
        }
    }
}

function Get-ZEPValuesByPropertyPath {
    param (
        [Parameter(Mandatory = $false)]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$PathSegments,
        [Parameter(Mandatory = $false)][int]$SegmentIndex = 0
    )

    if ($null -eq $InputObject) {
        return @()
    }

    if ($SegmentIndex -ge $PathSegments.Length) {
        return @($InputObject)
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $values = [System.Collections.Generic.List[object]]::new()
        foreach ($entry in $InputObject) {
            foreach ($value in (Get-ZEPValuesByPropertyPath -InputObject $entry -PathSegments $PathSegments -SegmentIndex $SegmentIndex)) {
                $values.Add($value)
            }
        }
        return @($values)
    }

    $segment = $PathSegments[$SegmentIndex]
    $nextValue = $null

    if ($InputObject -is [System.Collections.IDictionary]) {
        if (-not $InputObject.Contains($segment)) {
            return @()
        }
        $nextValue = $InputObject[$segment]
    }
    else {
        $property = $InputObject.PSObject.Properties | Where-Object { $_.Name -ieq $segment } | Select-Object -First 1
        if (-not $property) {
            return @()
        }
        $nextValue = $property.Value
    }

    return Get-ZEPValuesByPropertyPath -InputObject $nextValue -PathSegments $PathSegments -SegmentIndex ($SegmentIndex + 1)
}

<#
.SYNOPSIS
Searches ZEP offers by property and substring.

.DESCRIPTION
Searches offers by a specific property path (for example title or status.name)
and performs case-insensitive partial-string matching on the property value.
The allowed searchable properties are discovered from the last sample entries
in the local offers cache (default: last 1000 cache records). If the offers
cache is stale (TTL exceeded or total-count mismatch), the full offers cache is
rebuilt before searching.

.PARAMETER PropertyName
Property name or path to search, for example title, name, or status.name.

.PARAMETER SearchTerm
Substring to search for inside the selected property value.

.PARAMETER ReferenceSampleSize
Number of latest cache records used to determine allowed searchable properties.
Default is 1000.

.PARAMETER MaxResults
Maximum number of matching offers to return. Default is 500.

.PARAMETER CacheTtlMinutes
Maximum cache age in minutes before refresh. Default is module configuration.

.PARAMETER ValidateCache
Controls whether cached total count is validated against API before reuse.
Default is $true.

.PARAMETER SkipPropertyValidation
Skips pre-validation of PropertyName against reference cache records.
Use this for maximum performance when the property path is known to be valid.

.PARAMETER RefreshPageSize
Page size used when rebuilding the full offers cache.

.PARAMETER CachePath
Path to the offers cache JSON file.

.EXAMPLE
Search-ZEPOffers -PropertyName 'title' -SearchTerm 'test'

.EXAMPLE
Search-ZEPOffers -PropertyName 'status.name' -SearchTerm 'aktiv'

.EXAMPLE
Search-ZEPOffers -PropertyName 'name' -SearchTerm 'berlin' -MaxResults 100

.EXAMPLE
Search-ZEPOffers -PropertyName 'title' -SearchTerm 'test' -ReferenceSampleSize 1000

.EXAMPLE
Search-ZEPOffers -PropertyName 'title' -SearchTerm 'test' -CacheTtlMinutes 60 -ValidateCache

.EXAMPLE
Search-ZEPOffers -PropertyName 'title' -SearchTerm 'test' -ValidateCache:$false -SkipPropertyValidation
#>
function Search-ZEPOffers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [Parameter(Mandatory = $true)][string]$SearchTerm,
        [Parameter(Mandatory = $false)][int]$ReferenceSampleSize = 1000,
        [Parameter(Mandatory = $false)][int]$MaxResults = 500,
        [Parameter(Mandatory = $false)][int]$CacheTtlMinutes = $script:ZEPSettings.DefaultCacheTtlMinutes,
        [Parameter(Mandatory = $false)][bool]$ValidateCache = $true,
        [Parameter(Mandatory = $false)][switch]$SkipPropertyValidation,
        [Parameter(Mandatory = $false)][int]$RefreshPageSize = $script:ZEPSettings.DefaultPageSize,
        [Parameter(Mandatory = $false)][string]$CachePath = $script:ZEP_offer_cache_path
    )

    if ([string]::IsNullOrWhiteSpace($PropertyName)) {
        throw 'PropertyName darf nicht leer sein.'
    }

    if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
        throw 'SearchTerm darf nicht leer sein.'
    }

    $shouldValidateCache = [bool]$ValidateCache

    $cache = Read-ZEPOfferCache -Path $CachePath
    $cacheNeedsRefresh = $false

    if (-not $cache -or -not $cache.items) {
        $cacheNeedsRefresh = $true
        Write-Host 'Offers-Cache nicht vorhanden oder leer. Vollständiger Cache wird aufgebaut.' -ForegroundColor Yellow
    }
    elseif (Test-ZEPCacheAgeExceeded -CacheObject $cache -CacheTtlMinutes $CacheTtlMinutes) {
        $cacheNeedsRefresh = $true
        Write-Host "Offers-Cache älter als $CacheTtlMinutes Minuten. Vollständiger Cache wird aufgebaut." -ForegroundColor Yellow
    }
    elseif ($shouldValidateCache) {
        $cacheNeedsRefresh = Test-ZEPOffersCacheNeedsRefresh -CacheObject $cache -CacheTtlMinutes $CacheTtlMinutes
        if ($cacheNeedsRefresh) {
            Write-Host 'Offers-Cache ist inhaltlich veraltet (z. B. geänderte Gesamtanzahl). Vollständiger Cache wird aufgebaut.' -ForegroundColor Yellow
        }
    }

    if ($cacheNeedsRefresh) {
        # Forces a full fetch and cache rewrite before executing the search.
        Get-ZEPOffers -Refresh -PageSize $RefreshPageSize -CacheTtlMinutes $CacheTtlMinutes -CachePath $CachePath > $null
        $cache = Read-ZEPOfferCache -Path $CachePath
        if (-not $cache -or -not $cache.items) {
            throw "Offers-Cache konnte nicht aufgebaut werden: $CachePath"
        }
    }

    $allCacheItems = @($cache.items)
    if ($allCacheItems.Count -eq 0) {
        return @()
    }

    $segments = $PropertyName -split '\\.'
    if (-not $SkipPropertyValidation) {
        $sampleSize = [Math]::Max(1, [Math]::Min($ReferenceSampleSize, $allCacheItems.Count))
        $referenceItems = @($allCacheItems | Select-Object -Last $sampleSize)

        $propertyExists = $false
        foreach ($entry in $referenceItems) {
            if (-not $entry -or -not $entry.data) {
                continue
            }

            $referenceValues = Get-ZEPValuesByPropertyPath -InputObject $entry.data -PathSegments $segments
            if ($referenceValues -and @($referenceValues).Count -gt 0) {
                $propertyExists = $true
                break
            }
        }

        if (-not $propertyExists) {
            # Build deep property suggestions only on invalid input to avoid expensive traversal on normal searches.
            $propertySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($entry in $referenceItems) {
                if ($entry -and $entry.data) {
                    Get-ZEPPropertyPathsFromObject -InputObject $entry.data -PropertySet $propertySet
                }
            }

            $matchingProperties = @($propertySet | Where-Object { $_ -like "*$PropertyName*" } | Sort-Object)
            $preview = @($matchingProperties | Select-Object -First 20)
            $previewText = if ($preview.Count -gt 0) { ($preview -join ', ') } else { 'keine aehnlichen Felder gefunden' }
            throw "Property '$PropertyName' ist in den letzten $sampleSize Cache-Datensaetzen nicht vorhanden. Vorschlaege: $previewText"
        }
    }

    $result = [System.Collections.Generic.List[object]]::new()
    $max = [Math]::Max(1, $MaxResults)
    $isSingleSegment = $segments.Length -eq 1
    $singleSegmentName = if ($isSingleSegment) { $segments[0] } else { $null }

    foreach ($entry in $allCacheItems) {
        if (-not $entry -or -not $entry.data) {
            continue
        }

        if ($isSingleSegment) {
            $singleValue = $null
            if ($entry.data -is [System.Collections.IDictionary]) {
                if ($entry.data.Contains($singleSegmentName)) {
                    $singleValue = $entry.data[$singleSegmentName]
                }
            }
            else {
                $singleValue = $entry.data.$singleSegmentName
            }

            if ($singleValue -is [System.Collections.IEnumerable] -and -not ($singleValue -is [string])) {
                $values = @($singleValue)
            }
            elseif ($null -ne $singleValue) {
                $values = @($singleValue)
            }
            else {
                $values = @()
            }
        }
        else {
            $values = Get-ZEPValuesByPropertyPath -InputObject $entry.data -PathSegments $segments
        }

        if (-not $values -or @($values).Count -eq 0) {
            continue
        }

        $isMatch = $false
        foreach ($value in $values) {
            if ($null -eq $value) {
                continue
            }

            $candidate = [string]$value
            if ($candidate.IndexOf($SearchTerm, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $isMatch = $true
                break
            }
        }

        if ($isMatch) {
            $result.Add($entry.data)
            if ($result.Count -ge $max) {
                break
            }
        }
    }

    return @($result)
}

function Get-ZEPTimeZoneInfo {
    param (
        [Parameter(Mandatory = $false)][string]$Id = $script:ZEPSettings.ServerTimeZone
    )

    if (-not $PSBoundParameters.ContainsKey('Id') -and $script:ZEPTimeZoneInfo) {
        return $script:ZEPTimeZoneInfo
    }

    # Windows PowerShell 5.1 knows only Windows IDs, Linux (Azure Functions) only IANA IDs.
    $alternativeIds = @{
        'Europe/Berlin' = 'W. Europe Standard Time'
        'W. Europe Standard Time' = 'Europe/Berlin'
    }

    $timeZone = $null
    foreach ($candidate in @($Id, $alternativeIds[$Id]) | Where-Object { $_ }) {
        try {
            $timeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById($candidate)
            break
        }
        catch {
            continue
        }
    }

    if (-not $timeZone) {
        throw "Zeitzone '$Id' wurde auf diesem System nicht gefunden."
    }

    if (-not $PSBoundParameters.ContainsKey('Id')) {
        $script:ZEPTimeZoneInfo = $timeZone
    }

    return $timeZone
}

<#
.SYNOPSIS
Converts a ZEP timestamp into a correct UTC [datetime].

.DESCRIPTION
ZEP returns local wall-clock times (Europe/Berlin) but marks them with 'Z' as if they were UTC.
PowerShell 7 therefore parses them into a wrong UTC [datetime], Windows PowerShell 5.1 returns a string.
This function accepts both, takes the wall-clock time as sent by ZEP and converts it from the
configured ServerTimeZone into UTC (Kind = Utc).

Use -AsDate for date-only fields such as valid_until: the calendar date is kept unchanged and
returned as midnight UTC.

.EXAMPLE
ConvertFrom-ZEPDateTime -Value '2026-10-07T01:24:35.000000Z'
# -> 06.10.2026 23:24:35 (Kind Utc)

.EXAMPLE
ConvertFrom-ZEPDateTime -Value '2026-12-31T00:00:00.000000Z' -AsDate
# -> 31.12.2026 00:00:00 (Kind Utc)

.EXAMPLE
(Get-ZEPOffer -Id 18380).status_since | ConvertFrom-ZEPDateTime
#>
function ConvertFrom-ZEPDateTime {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, ValueFromPipeline = $true)]$Value,
        [Parameter(Mandatory = $false)][switch]$AsDate
    )

    process {
        if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) {
            return $null
        }

        if ($Value -is [datetimeoffset]) {
            $Value = $Value.UtcDateTime
        }

        if ($Value -is [datetime]) {
            # PowerShell 7 parses '...Z' as Utc (wall clock unchanged) or Local (shifted); recover the wall clock sent by ZEP.
            $wallClock = if ($Value.Kind -eq [DateTimeKind]::Local) { $Value.ToUniversalTime() } else { $Value }
        }
        else {
            $text = ([string]$Value).Trim() -replace '(Z|[+-]\d{2}:?\d{2})$', ''
            $text = $text -replace '(\.\d{7})\d+$', '$1'
            try {
                $wallClock = [datetime]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None)
            }
            catch {
                throw "Ungueltiger ZEP-Zeitstempel: '$Value'"
            }
        }

        $wallClock = [datetime]::SpecifyKind($wallClock, [DateTimeKind]::Unspecified)

        if ($AsDate.IsPresent) {
            return [datetime]::SpecifyKind($wallClock.Date, [DateTimeKind]::Utc)
        }

        $timeZone = Get-ZEPTimeZoneInfo
        if ($timeZone.IsInvalidTime($wallClock)) {
            # Wall-clock times inside the spring-forward gap do not exist; move them past the gap.
            $wallClock = $wallClock.AddHours(1)
        }

        return [System.TimeZoneInfo]::ConvertTimeToUtc($wallClock, $timeZone)
    }
}

<#
.SYNOPSIS
Converts a raw ZEP offer into an offer with reliable PowerShell types.

.DESCRIPTION
Keeps all original property names and converts the values:
- id, status.id and other numeric IDs/counters to [int]
- customer_id and customer_contact_id to [string] (customer numbers are strings in ZEP)
- status_since to a correct UTC [datetime] (see ConvertFrom-ZEPDateTime)
- valid_until, order_date, realization_from, realization_until to dates (midnight UTC)

.EXAMPLE
Get-ZEPOffer -Id 18380 -UseCache:$false | ConvertTo-ZEPTypedOffer

.EXAMPLE
Get-ZEPOffersById -Id 18380 -Raw | ConvertTo-ZEPTypedOffer | Select-Object id, status_since, valid_until
#>
function ConvertTo-ZEPTypedOffer {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]$Offer
    )

    process {
        if ($null -eq $Offer) {
            return
        }

        $typed = ConvertTo-ZEPObject -InputObject $Offer

        foreach ($name in @('id', 'department_id', 'project_id', 'ticket_id', 'default_validity', 'items_count', 'order_probability_percent', 'is_net')) {
            if ($typed.PSObject.Properties.Name -contains $name -and $null -ne $typed.$name -and [string]$typed.$name -ne '') {
                $typed.$name = [int]$typed.$name
            }
        }

        foreach ($name in @('customer_id', 'customer_contact_id')) {
            if ($typed.PSObject.Properties.Name -contains $name -and $null -ne $typed.$name) {
                $typed.$name = [string]$typed.$name
            }
        }

        if ($typed.PSObject.Properties.Name -contains 'status_since') {
            $typed.status_since = ConvertFrom-ZEPDateTime -Value $typed.status_since
        }

        foreach ($name in @('valid_until', 'order_date', 'realization_from', 'realization_until')) {
            if ($typed.PSObject.Properties.Name -contains $name) {
                $typed.$name = ConvertFrom-ZEPDateTime -Value $typed.$name -AsDate
            }
        }

        if ($typed.status -and $typed.status.PSObject.Properties.Name -contains 'id' -and $null -ne $typed.status.id) {
            $typed.status.id = [int]$typed.status.id
        }

        return $typed
    }
}

<#
.SYNOPSIS
Gets the newest ZEP offers without downloading all offers.

.DESCRIPTION
The ZEP offers list ignores sorting and filter parameters and is always sorted by ascending id.
This function therefore reads the list backwards from the last page until it reaches -AfterId
(or has collected -Last offers). Typically 1-3 API calls instead of a full download.
Results are typed via ConvertTo-ZEPTypedOffer (use -Raw for the unchanged API objects)
and returned in ascending id order.

.PARAMETER AfterId
Returns all offers with an id greater than this value (watermark).

.PARAMETER Last
Returns the newest N offers.

.PARAMETER MaxPages
Safety limit for pages read per call; a warning is written when it is reached.

.EXAMPLE
Get-ZEPRecentOffers -AfterId 18377

.EXAMPLE
Get-ZEPRecentOffers -Last 20 | Select-Object id, title, @{ n = 'status'; e = { $_.status.name } }, status_since

.EXAMPLE
# Neue Angebote eines Bearbeiters seit dem Wasserzeichen.
Get-ZEPRecentOffers -AfterId 18377 | Where-Object processor_username -eq 'm.mustermann'
#>
function Get-ZEPRecentOffers {
    [CmdletBinding(DefaultParameterSetName = 'Last')]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'AfterId')][int]$AfterId,
        [Parameter(Mandatory = $true, ParameterSetName = 'Last')][ValidateRange(1, 2000)][int]$Last,
        [Parameter(Mandatory = $false)][ValidateRange(1, 200)][int]$PageSize = $script:ZEPSettings.DefaultPageSize,
        [Parameter(Mandatory = $false)][ValidateRange(1, 1000)][int]$MaxPages = 20,
        [Parameter(Mandatory = $false)][switch]$Raw
    )

    $probe = Invoke-ZEPRest -Path 'offers' -Query @{ limit = 1; page = 1 }
    $total = Get-ZEPResponseTotal -Payload $probe
    if (-not $total) {
        return @()
    }

    $collected = @{}
    $page = [int][Math]::Ceiling($total / $PageSize)
    $pagesRead = 0
    $isFirstPage = $true

    while ($page -ge 1) {
        if ($pagesRead -ge $MaxPages) {
            Write-Warning "Get-ZEPRecentOffers: MaxPages ($MaxPages) erreicht, Ergebnis ist moeglicherweise unvollstaendig."
            break
        }

        $response = Invoke-ZEPRest -Path 'offers' -Query @{ limit = $PageSize; page = $page }
        $pagesRead++
        $pageItems = @(Get-ZEPResponseItems -Payload $response)

        foreach ($item in $pageItems) {
            if ($item -and $null -ne $item.id) {
                $collected[[int]$item.id] = $item
            }
        }

        # Offers created between the probe and this call may have opened a new last page.
        if ($isFirstPage -and $response.meta -and [int]$response.meta.last_page -gt $page) {
            $page = [int]$response.meta.last_page
            $isFirstPage = $false
            continue
        }
        $isFirstPage = $false

        if ($pageItems.Count -eq 0) {
            $page--
            continue
        }

        $minId = ($pageItems | ForEach-Object { [int]$_.id } | Measure-Object -Minimum).Minimum
        if ($PSCmdlet.ParameterSetName -eq 'AfterId' -and $minId -le $AfterId) {
            break
        }

        if ($PSCmdlet.ParameterSetName -eq 'Last' -and $collected.Count -ge $Last) {
            break
        }

        $page--
    }

    $ids = @($collected.Keys | Sort-Object)
    if ($PSCmdlet.ParameterSetName -eq 'AfterId') {
        $ids = @($ids | Where-Object { $_ -gt $AfterId })
    }
    else {
        $ids = @($ids | Select-Object -Last $Last)
    }

    foreach ($id in $ids) {
        if ($Raw.IsPresent) {
            ConvertTo-ZEPObject -InputObject $collected[$id]
        }
        else {
            ConvertTo-ZEPTypedOffer -Offer $collected[$id]
        }
    }
}

<#
.SYNOPSIS
Gets several ZEP offers by id in as few API calls as possible.

.DESCRIPTION
Uses the id[] filter of the offers list (up to 100 ids per call). Offers that do not exist
(e.g. deleted) are simply missing from the result. Results are typed via
ConvertTo-ZEPTypedOffer (use -Raw for the unchanged API objects), sorted by id.

.EXAMPLE
Get-ZEPOffersById -Id 18380

.EXAMPLE
Get-ZEPOffersById -Id 18368, 18380, 18381 | Select-Object id, @{ n = 'status'; e = { $_.status.id } }, status_since
#>
function Get-ZEPOffersById {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][int[]]$Id,
        [Parameter(Mandatory = $false)][switch]$Raw
    )

    begin {
        $allIds = [System.Collections.Generic.List[int]]::new()
    }

    process {
        foreach ($value in $Id) {
            if (-not $allIds.Contains($value)) {
                $allIds.Add($value)
            }
        }
    }

    end {
        $chunkSize = 100
        $records = [System.Collections.Generic.List[object]]::new()

        for ($offset = 0; $offset -lt $allIds.Count; $offset += $chunkSize) {
            $chunk = @($allIds.GetRange($offset, [Math]::Min($chunkSize, $allIds.Count - $offset)))
            $response = Invoke-ZEPRest -Path 'offers' -Query @{ 'id[]' = $chunk; limit = $chunk.Count; page = 1 }
            foreach ($item in @(Get-ZEPResponseItems -Payload $response)) {
                # Guard against an API that ignores the filter: only return requested ids.
                if ($item -and $chunk -contains [int]$item.id) {
                    $records.Add($item)
                }
            }
        }

        foreach ($record in ($records | Sort-Object { [int]$_.id })) {
            if ($Raw.IsPresent) {
                ConvertTo-ZEPObject -InputObject $record
            }
            else {
                ConvertTo-ZEPTypedOffer -Offer $record
            }
        }
    }
}

<#
.SYNOPSIS
Gets a ZEP employee by username.

.DESCRIPTION
Reads one employee (e.g. to resolve a username to the email address). Results are cached
per session; use -Refresh to bypass the cache. Returns $null when the username does not exist.

.EXAMPLE
Get-ZEPEmployee -Username 'm.mustermann'

.EXAMPLE
(Get-ZEPEmployee -Username 'm.mustermann').email
#>
function Get-ZEPEmployee {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$Username,
        [Parameter(Mandatory = $false)][switch]$Refresh
    )

    process {
        if (-not $script:ZEPEmployeeCache) {
            $script:ZEPEmployeeCache = @{}
        }

        $key = $Username.Trim().ToLowerInvariant()
        if (-not $Refresh.IsPresent -and $script:ZEPEmployeeCache.ContainsKey($key)) {
            return $script:ZEPEmployeeCache[$key]
        }

        $employee = $null
        try {
            $response = Invoke-ZEPRest -Path "employees/$([uri]::EscapeDataString($Username.Trim()))"
            $employee = ConvertTo-ZEPObject -InputObject (@(Get-ZEPResponseItems -Payload $response) | Select-Object -First 1)
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($statusCode -ne 404) {
                throw
            }
        }

        $script:ZEPEmployeeCache[$key] = $employee
        return $employee
    }
}

Export-ModuleMember Get-ZEPConfiguration
Export-ModuleMember Set-ZEPConfiguration
Export-ModuleMember Invoke-ZEPRest
Export-ModuleMember Get-ZEPOffer
Export-ModuleMember Get-ZEPOffers
Export-ModuleMember Get-ZEPOfferItems
Export-ModuleMember Test-ZEPOffersCacheNeedsRefresh
Export-ModuleMember Search-ZEPOffers
Export-ModuleMember ConvertFrom-ZEPDateTime
Export-ModuleMember ConvertTo-ZEPTypedOffer
Export-ModuleMember Get-ZEPRecentOffers
Export-ModuleMember Get-ZEPOffersById
Export-ModuleMember Get-ZEPEmployee
