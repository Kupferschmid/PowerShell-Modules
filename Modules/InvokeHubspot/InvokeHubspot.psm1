# InvokeHubspot
# Private App Bearer token support for HubSpot CRM Deals API.
# Supports local sessions via Windows Credential Manager and Azure Automation via Get-AutomationPSCredential.

# Version 0.4.3 07.10.2026 by Klaus Kupferschmid (tempero.it GmbH & hhpberlin GmbH)

#Requires -Modules @{ ModuleName = 'BetterCredentials'; ModuleVersion = '4.5' }

# Windows PowerShell 5.1 does not load System.Web (HttpUtility) by default.
Add-Type -AssemblyName System.Web

$script:HubspotSettings = [ordered]@{
    ServiceUserName = 'Hubspot_Deals'
    BaseUri = 'https://api.hubapi.com'
    # Date-based HubSpot API version (/crm/{api}/YYYY-MM/...). Legacy v1-v4 APIs are unsupported from September 2027.
    ApiVersion = '2026-09'
    # Optional override for the deals endpoint. Empty = derived from ApiVersion.
    DealsPath = $null
    TokenTarget = 'Hubspot_Deals_BearerToken'
    DefaultPageSize = 100
    DefaultRetryCount = 2
    DefaultThrottleDelaySeconds = 0.2
}

# HubSpot-defined association type IDs (deal -> company).
$script:HubspotDealToPrimaryCompanyTypeId = 5
$script:HubspotDealToCompanyTypeId = 341

<#
.SYNOPSIS
Shows the active InvokeHubspot module configuration.

.DESCRIPTION
Returns the current runtime settings such as API base URI, API version, effective deals path,
default paging size, retry behavior, and credential target name.

.EXAMPLE
Get-HubspotConfiguration

.EXAMPLE
Get-HubspotConfiguration | Format-List *

.EXAMPLE
# Troubleshooting: Aktive Runtime-Werte pruefen, wenn Paging/Retry unerwartet ist.
Get-HubspotConfiguration | Select-Object DefaultPageSize, DefaultRetryCount, DefaultThrottleDelaySeconds
#>
function Get-HubspotConfiguration {
    [CmdletBinding()]
    param ()

    return [pscustomobject]@{
        ServiceUserName = $script:HubspotSettings.ServiceUserName
        BaseUri = $script:HubspotSettings.BaseUri
        ApiVersion = $script:HubspotSettings.ApiVersion
        DealsPath = Get-HubspotDealsPath
        TokenTarget = $script:HubspotSettings.TokenTarget
        DefaultPageSize = $script:HubspotSettings.DefaultPageSize
        DefaultRetryCount = $script:HubspotSettings.DefaultRetryCount
        DefaultThrottleDelaySeconds = $script:HubspotSettings.DefaultThrottleDelaySeconds
    }
}

<#
.SYNOPSIS
Updates InvokeHubspot module configuration values.

.DESCRIPTION
Changes selected runtime settings for the current session, for example default page size,
base URI, API version (date-based, e.g. '2026-09'), retry count, and credential target name.

.EXAMPLE
Set-HubspotConfiguration -DefaultPageSize 200

.EXAMPLE
Set-HubspotConfiguration -DefaultRetryCount 3 -DefaultThrottleDelaySeconds 0.5

.EXAMPLE
Set-HubspotConfiguration -ApiVersion '2026-09'

.EXAMPLE
# DealsPath explizit ueberschreiben bzw. mit '' wieder aus ApiVersion ableiten lassen.
Set-HubspotConfiguration -BaseUri 'https://api.hubapi.com' -DealsPath 'crm/objects/2026-09/deals'
Set-HubspotConfiguration -DealsPath ''

.EXAMPLE
Set-HubspotConfiguration -ServiceUserName 'Hubspot_Deals' -TokenTarget 'Hubspot_Deals_BearerToken'

.EXAMPLE
# Troubleshooting: Nach Experimenten auf stabile Defaults zuruecksetzen.
Set-HubspotConfiguration -DefaultPageSize 100 -DefaultRetryCount 2 -DefaultThrottleDelaySeconds 0.2
#>
function Set-HubspotConfiguration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)][string]$ServiceUserName,
        [Parameter(Mandatory = $false)][string]$BaseUri,
        [Parameter(Mandatory = $false)][ValidatePattern('^\d{4}-\d{2}(-beta)?$')][string]$ApiVersion,
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$DealsPath,
        [Parameter(Mandatory = $false)][string]$TokenTarget,
        [Parameter(Mandatory = $false)][int]$DefaultPageSize,
        [Parameter(Mandatory = $false)][int]$DefaultRetryCount,
        [Parameter(Mandatory = $false)][double]$DefaultThrottleDelaySeconds
    )

    if ($PSBoundParameters.ContainsKey('ServiceUserName')) {
        $script:HubspotSettings.ServiceUserName = $ServiceUserName
    }

    if ($PSBoundParameters.ContainsKey('BaseUri')) {
        $script:HubspotSettings.BaseUri = $BaseUri.TrimEnd('/')
    }

    if ($PSBoundParameters.ContainsKey('ApiVersion')) {
        $script:HubspotSettings.ApiVersion = $ApiVersion
        # Cached metadata may differ between API versions.
        $script:HubspotPropertyDefinitionsByObjectType = $null
        $script:HubspotDealPipelineDefinitionsById = $null
        $script:HubspotDealPipelineDefinitionsLoaded = $false
    }

    if ($PSBoundParameters.ContainsKey('DealsPath')) {
        $script:HubspotSettings.DealsPath = $DealsPath.Trim('/ ')
    }

    if ($PSBoundParameters.ContainsKey('TokenTarget')) {
        $script:HubspotSettings.TokenTarget = $TokenTarget
    }

    if ($PSBoundParameters.ContainsKey('DefaultPageSize')) {
        $script:HubspotSettings.DefaultPageSize = [Math]::Max(1, [Math]::Min(200, $DefaultPageSize))
    }

    if ($PSBoundParameters.ContainsKey('DefaultRetryCount')) {
        $script:HubspotSettings.DefaultRetryCount = [Math]::Max(0, $DefaultRetryCount)
    }

    if ($PSBoundParameters.ContainsKey('DefaultThrottleDelaySeconds')) {
        $script:HubspotSettings.DefaultThrottleDelaySeconds = [Math]::Max(0.0, $DefaultThrottleDelaySeconds)
    }

    return Get-HubspotConfiguration
}

function Get-HubspotCrmPath {
    param (
        [Parameter(Mandatory = $true)][ValidateSet('objects', 'properties', 'pipelines', 'owners')][string]$Api,
        [Parameter(Mandatory = $false)][string]$Resource
    )

    $path = "crm/$Api/$($script:HubspotSettings.ApiVersion)"
    if (-not [string]::IsNullOrWhiteSpace($Resource)) {
        $path = "$path/$($Resource.Trim('/'))"
    }

    return $path
}

function Get-HubspotDealsPath {
    if (-not [string]::IsNullOrWhiteSpace($script:HubspotSettings.DealsPath)) {
        return $script:HubspotSettings.DealsPath
    }

    return Get-HubspotCrmPath -Api 'objects' -Resource 'deals'
}

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

function Test-HubspotAuthenticationFailure {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]$ErrorRecord,
        [Parameter(Mandatory = $false)][int]$StatusCode
    )

    $errorMessage = [string]$ErrorRecord.Exception.Message
    $errorDetailsMessage = [string]$ErrorRecord.ErrorDetails.Message

    $hasInvalidAuthIndicator = (
        $errorMessage -match 'INVALID_AUTHENTICATION' -or
        $errorMessage -match 'Authentication credentials not found' -or
        $errorMessage -match 'OAuth 2\.0 authentication' -or
        $errorDetailsMessage -match 'INVALID_AUTHENTICATION' -or
        $errorDetailsMessage -match 'Authentication credentials not found' -or
        $errorDetailsMessage -match 'OAuth 2\.0 authentication'
    )

    if ($hasInvalidAuthIndicator) {
        return $true
    }

    if ($StatusCode -in @(401, 403)) {
        return $true
    }

    return $false
}

function Test-HubspotMissingScopeFailure {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]$ErrorRecord,
        [Parameter(Mandatory = $false)][string]$ExpectedScope,
        [Parameter(Mandatory = $false)][string]$ObjectHint
    )

    $statusCode = $null
    if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
        $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode
    }

    $errorMessage = [string]$ErrorRecord.Exception.Message
    $errorDetailsMessage = [string]$ErrorRecord.ErrorDetails.Message
    $combinedMessage = "$errorMessage $errorDetailsMessage"

    $hasScopeIndicator = (
        $combinedMessage -match 'insufficient scopes|missing scopes|required scopes|forbidden|insufficient permissions|permission denied'
    )

    if (-not $hasScopeIndicator -and $statusCode -ne 403) {
        return $false
    }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedScope) -and $combinedMessage -match [Regex]::Escape($ExpectedScope)) {
        return $true
    }

    if (-not [string]::IsNullOrWhiteSpace($ObjectHint) -and $combinedMessage -match [Regex]::Escape($ObjectHint)) {
        return $true
    }

    if ($statusCode -eq 403) {
        return $true
    }

    return $false
}

function Write-HubspotMissingScopeWarning {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$Scope,
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [Parameter(Mandatory = $true)][string]$ObjectLabel
    )

    if (-not $script:HubspotMissingScopeWarningsShown) {
        $script:HubspotMissingScopeWarningsShown = @{}
    }

    $warningKey = "$Scope|$PropertyName"
    if ($script:HubspotMissingScopeWarningsShown.ContainsKey($warningKey)) {
        return
    }

    Write-Warning "DisplayValue fuer Property '$PropertyName' kann nicht aufgeloest werden. Fuer $ObjectLabel wird Scope $Scope benoetigt. Bis dahin bleibt DisplayValue die interne ID."
    $script:HubspotMissingScopeWarningsShown[$warningKey] = $true
}

function Invoke-HubspotRequest {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$ResourcePath,
        [Parameter(Mandatory = $false)][string]$Method = 'GET',
        [Parameter(Mandatory = $false)][hashtable]$Query,
        [Parameter(Mandatory = $false)]$Body,
        [Parameter(Mandatory = $false)][string]$AccessToken,
        [Parameter(Mandatory = $false)][int]$RetryCount = $script:HubspotSettings.DefaultRetryCount,
        [Parameter(Mandatory = $false)][double]$ThrottleDelaySeconds = $script:HubspotSettings.DefaultThrottleDelaySeconds
    )

    $token = Get-HubspotBearerToken -AccessToken $AccessToken
    $authRetryTriggered = $false
    $headers = @{
        Accept = 'application/json'
        Authorization = "Bearer $token"
    }

    $resourceUri = $script:HubspotSettings.BaseUri.TrimEnd('/')
    if ($ResourcePath) {
        $resourceUri = "$resourceUri/$($ResourcePath.TrimStart('/'))"
    }

    $uriBuilder = [System.UriBuilder]::new($resourceUri)
    if ($Query -and $Query.Count -gt 0) {
        $queryCollection = [System.Web.HttpUtility]::ParseQueryString([string]::Empty)
        foreach ($key in $Query.Keys) {
            $value = $Query[$key]
            if ($null -eq $value) {
                continue
            }

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

            $isAuthenticationFailure = Test-HubspotAuthenticationFailure -ErrorRecord $_ -StatusCode $statusCode
            if ($isAuthenticationFailure -and [string]::IsNullOrWhiteSpace($AccessToken) -and -not $authRetryTriggered) {
                try {
                    $token = Get-HubspotBearerToken -ForcePrompt
                }
                catch {
                    throw
                }

                $headers.Authorization = "Bearer $token"
                $authRetryTriggered = $true
                continue
            }

            if ($attempt -lt $RetryCount -and ($statusCode -in @(429, 500, 502, 503, 504))) {
                $backoffSeconds = [Math]::Min(20, [Math]::Pow(2, $attempt + 1))
                Start-Sleep -Seconds $backoffSeconds
                continue
            }

            throw
        }
    }
}

function Get-HubspotPropertyDefinitions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)][string]$ObjectType = 'deals',
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    if (-not $script:HubspotPropertyDefinitionsByObjectType) {
        $script:HubspotPropertyDefinitionsByObjectType = @{}
    }

    if ($script:HubspotPropertyDefinitionsByObjectType.ContainsKey($ObjectType)) {
        return $script:HubspotPropertyDefinitionsByObjectType[$ObjectType]
    }

    $response = Invoke-HubspotRequest -ResourcePath (Get-HubspotCrmPath -Api 'properties' -Resource $ObjectType) -AccessToken $AccessToken
    $definitionMap = @{}

    foreach ($definition in @($response.results)) {
        if ($definition -and $definition.name) {
            $definitionMap[$definition.name] = $definition
        }
    }

    $script:HubspotPropertyDefinitionsByObjectType[$ObjectType] = $definitionMap
    return $definitionMap
}

function Get-HubspotDealStageDefinition {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)][string]$PipelineId,
        [Parameter(Mandatory = $false)][string]$StageId,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    if ([string]::IsNullOrWhiteSpace($StageId)) {
        return $null
    }

    if (-not $script:HubspotDealPipelineDefinitionsById) {
        $script:HubspotDealPipelineDefinitionsById = @{ }
    }

    if (-not $script:HubspotDealPipelineDefinitionsLoaded) {
        $response = Invoke-HubspotRequest -ResourcePath (Get-HubspotCrmPath -Api 'pipelines' -Resource 'deals') -AccessToken $AccessToken

        foreach ($pipeline in @($response.results)) {
            if (-not $pipeline -or -not $pipeline.id) {
                continue
            }

            $stageMap = @{}
            foreach ($stage in @($pipeline.stages)) {
                if ($stage -and $stage.id) {
                    $stageMap[[string]$stage.id] = [pscustomobject]@{
                        Id = [string]$stage.id
                        Label = [string]$stage.label
                        PipelineId = [string]$pipeline.id
                        PipelineLabel = [string]$pipeline.label
                    }
                }
            }

            $script:HubspotDealPipelineDefinitionsById[[string]$pipeline.id] = [pscustomobject]@{
                Id = [string]$pipeline.id
                Label = [string]$pipeline.label
                Stages = $stageMap
            }
        }

        $script:HubspotDealPipelineDefinitionsLoaded = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($PipelineId) -and $script:HubspotDealPipelineDefinitionsById.ContainsKey($PipelineId)) {
        $pipelineDefinition = $script:HubspotDealPipelineDefinitionsById[$PipelineId]
        if ($pipelineDefinition.Stages.ContainsKey($StageId)) {
            return $pipelineDefinition.Stages[$StageId]
        }
    }

    foreach ($pipelineDefinition in @($script:HubspotDealPipelineDefinitionsById.Values)) {
        if ($pipelineDefinition.Stages.ContainsKey($StageId)) {
            return $pipelineDefinition.Stages[$StageId]
        }
    }

    return $null
}

function Get-HubspotOwnerDisplayValue {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)][string]$OwnerId,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    if ([string]::IsNullOrWhiteSpace($OwnerId)) {
        return $OwnerId
    }

    if (-not $script:HubspotOwnerDisplayById) {
        $script:HubspotOwnerDisplayById = @{}
    }

    if ($script:HubspotOwnerDisplayById.ContainsKey($OwnerId)) {
        return $script:HubspotOwnerDisplayById[$OwnerId]
    }

    $displayValue = $null
    try {
        $owner = Invoke-HubspotRequest -ResourcePath (Get-HubspotCrmPath -Api 'owners' -Resource $OwnerId) -AccessToken $AccessToken
        if ($owner) {
            $fullName = ("{0} {1}" -f ([string]$owner.firstName), ([string]$owner.lastName)).Trim()
            if (-not [string]::IsNullOrWhiteSpace($fullName)) {
                $displayValue = $fullName
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$owner.email)) {
                $displayValue = [string]$owner.email
            }
        }
    }
    catch {
        if (Test-HubspotMissingScopeFailure -ErrorRecord $PSItem -ExpectedScope 'crm.objects.owners.read' -ObjectHint 'owners') {
            Write-HubspotMissingScopeWarning -Scope 'crm.objects.owners.read' -PropertyName 'hubspot_owner_id' -ObjectLabel 'Owner-Aufloesung'
        }

        $displayValue = $null
    }

    if ([string]::IsNullOrWhiteSpace($displayValue)) {
        $displayValue = $OwnerId
    }

    $script:HubspotOwnerDisplayById[$OwnerId] = $displayValue
    return $displayValue
}

function Get-HubspotRelatedObjectDisplayValue {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [Parameter(Mandatory = $true)][string]$ObjectType,
        [Parameter(Mandatory = $false)][string]$ObjectId,
        [Parameter(Mandatory = $false)][string[]]$QueryProperties,
        [Parameter(Mandatory = $false)][string[]]$DisplayPropertyCandidates,
        [Parameter(Mandatory = $true)][string]$RequiredReadScope,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    if ([string]::IsNullOrWhiteSpace($ObjectId)) {
        return $ObjectId
    }

    if (-not $script:HubspotRelatedDisplayByTypeAndId) {
        $script:HubspotRelatedDisplayByTypeAndId = @{}
    }

    if (-not $script:HubspotRelatedDisplayByTypeAndId.ContainsKey($ObjectType)) {
        $script:HubspotRelatedDisplayByTypeAndId[$ObjectType] = @{}
    }

    $objectCache = $script:HubspotRelatedDisplayByTypeAndId[$ObjectType]
    if ($objectCache.ContainsKey($ObjectId)) {
        return $objectCache[$ObjectId]
    }

    $displayValue = $null
    try {
        $query = @{}
        if ($QueryProperties -and $QueryProperties.Count -gt 0) {
            $query.properties = $QueryProperties
        }

        $record = Invoke-HubspotRequest -ResourcePath (Get-HubspotCrmPath -Api 'objects' -Resource "$ObjectType/$ObjectId") -Query $query -AccessToken $AccessToken
        $propertyBag = $record.properties

        if ($ObjectType -eq 'contacts' -and $propertyBag) {
            $fullName = ("{0} {1}" -f ([string]$propertyBag.firstname), ([string]$propertyBag.lastname)).Trim()
            if (-not [string]::IsNullOrWhiteSpace($fullName)) {
                $displayValue = $fullName
            }
        }

        if ([string]::IsNullOrWhiteSpace($displayValue) -and $propertyBag -and $DisplayPropertyCandidates) {
            foreach ($candidateProperty in $DisplayPropertyCandidates) {
                if ($propertyBag.PSObject.Properties.Name -contains $candidateProperty) {
                    $candidateValue = [string]$propertyBag.$candidateProperty
                    if (-not [string]::IsNullOrWhiteSpace($candidateValue)) {
                        $displayValue = $candidateValue
                        break
                    }
                }
            }
        }

        if ([string]::IsNullOrWhiteSpace($displayValue)) {
            foreach ($fallbackName in @('name', 'hs_name', 'title', 'subject', 'email')) {
                if ($propertyBag -and ($propertyBag.PSObject.Properties.Name -contains $fallbackName)) {
                    $candidateValue = [string]$propertyBag.$fallbackName
                    if (-not [string]::IsNullOrWhiteSpace($candidateValue)) {
                        $displayValue = $candidateValue
                        break
                    }
                }
            }
        }
    }
    catch {
        if (Test-HubspotMissingScopeFailure -ErrorRecord $PSItem -ExpectedScope $RequiredReadScope -ObjectHint $ObjectType) {
            Write-HubspotMissingScopeWarning -Scope $RequiredReadScope -PropertyName $PropertyName -ObjectLabel "$ObjectType-Aufloesung"
        }
    }

    if ([string]::IsNullOrWhiteSpace($displayValue)) {
        $displayValue = $ObjectId
    }

    $objectCache[$ObjectId] = $displayValue
    return $displayValue
}

function ConvertTo-HubspotWriteValue {
    param (
        [Parameter(Mandatory = $false)]$Value,
        [Parameter(Mandatory = $false)]$Definition
    )

    if ($null -eq $Value -or $Value -is [string]) {
        return $Value
    }

    $invariant = [System.Globalization.CultureInfo]::InvariantCulture

    if ($Value -is [datetimeoffset]) {
        $Value = $Value.UtcDateTime
    }

    if ($Value -is [datetime]) {
        # HubSpot date properties expect the calendar date (midnight UTC); datetime properties an ISO 8601 UTC timestamp.
        if ($Definition -and [string]$Definition.type -eq 'date') {
            return $Value.ToString('yyyy-MM-dd', $invariant)
        }

        return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ', $invariant)
    }

    if ($Value -is [bool]) {
        return $(if ($Value) { 'true' } else { 'false' })
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal] -or $Value -is [single]) {
        return $Value.ToString($invariant)
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        # Multi-select (checkbox) enumerations are written as semicolon-separated values.
        return (@($Value | ForEach-Object { ConvertTo-HubspotWriteValue -Value $_ -Definition $Definition }) -join ';')
    }

    return $Value
}

function Convert-HubspotDealPropertiesForWrite {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][hashtable]$Properties,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    $writeProperties = [ordered]@{}

    foreach ($entry in $Properties.GetEnumerator()) {
        if ([string]::IsNullOrWhiteSpace([string]$entry.Key)) {
            continue
        }

        if ($entry.Key -ieq 'dealstageId') {
            $writeProperties.dealstage = [string]$entry.Value
            continue
        }

        if ($entry.Key -ieq 'dealstageLabel') {
            continue
        }

        if ($entry.Key -ieq 'dealstage') {
            $stageValue = [string]$entry.Value
            $stageDefinition = Get-HubspotDealStageDefinition -StageId $stageValue -AccessToken $AccessToken

            if ($stageDefinition) {
                $writeProperties.dealstage = $stageDefinition.Id
                continue
            }

            $writeProperties.dealstage = $stageValue
            continue
        }

        $definition = $null
        if ($entry.Value -is [datetime] -or $entry.Value -is [datetimeoffset]) {
            $propertyDefinitions = Get-HubspotPropertyDefinitions -ObjectType 'deals' -AccessToken $AccessToken
            if ($propertyDefinitions -and $propertyDefinitions.ContainsKey([string]$entry.Key)) {
                $definition = $propertyDefinitions[[string]$entry.Key]
            }
        }

        $writeProperties[$entry.Key] = ConvertTo-HubspotWriteValue -Value $entry.Value -Definition $definition
    }

    return $writeProperties
}

function Convert-HubspotPropertyValue {
    param (
        [Parameter(Mandatory = $false)]$Value,
        [Parameter(Mandatory = $false)]$Definition
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $nestedObject = [ordered]@{}
        foreach ($key in $Value.Keys) {
            $nestedObject[$key] = Convert-HubspotPropertyValue -Value $Value[$key]
        }
        return [pscustomobject]$nestedObject
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $convertedItems = foreach ($item in $Value) {
            Convert-HubspotPropertyValue -Value $item
        }
        return @($convertedItems)
    }

    $propertyType = $null
    $fieldType = $null
    if ($Definition) {
        $propertyType = [string]$Definition.type
        $fieldType = [string]$Definition.fieldType
    }

    switch ($propertyType) {
        'bool' {
            try {
                return [System.Convert]::ToBoolean($Value)
            }
            catch {
                return $Value
            }
        }

        'number' {
            $numericValue = 0.0
            if ([double]::TryParse([string]$Value, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$numericValue)) {
                return $numericValue
            }

            return $Value
        }

        'date' {
            try {
                if ($Value -is [datetime]) {
                    $utcValue = if ($Value.Kind -eq [DateTimeKind]::Local) { $Value.ToUniversalTime() } else { $Value }
                    return [datetime]::SpecifyKind($utcValue.Date, [DateTimeKind]::Utc)
                }

                if ([string]$Value -match '^\d+$') {
                    return [datetimeoffset]::FromUnixTimeMilliseconds([int64]$Value).UtcDateTime.Date
                }

                return [datetime]::Parse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, ([System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)).Date
            }
            catch {
                return $Value
            }
        }

        'datetime' {
            try {
                # PowerShell 7 already converts ISO strings to [datetime]; keep the instant and mark it as UTC.
                if ($Value -is [datetime]) {
                    if ($Value.Kind -eq [DateTimeKind]::Local) {
                        return $Value.ToUniversalTime()
                    }

                    return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc)
                }

                if ([string]$Value -match '^\d+$') {
                    return [datetimeoffset]::FromUnixTimeMilliseconds([int64]$Value).UtcDateTime
                }

                return [datetime]::Parse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
            }
            catch {
                return $Value
            }
        }

        'enumeration' {
            if ($fieldType -eq 'checkbox' -or ([string]$Value -like '*;*')) {
                $selectedValues = @(
                    [string]$Value -split ';' |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                )

                if ($selectedValues.Count -gt 1) {
                    return $selectedValues
                }
            }

            return $Value
        }

        'json' {
            try {
                # ConvertFrom-Json -Depth exists only from PowerShell 6.2 on.
                if ($PSVersionTable.PSVersion -ge [version]'6.2') {
                    return [string]$Value | ConvertFrom-Json -Depth 20
                }

                return [string]$Value | ConvertFrom-Json
            }
            catch {
                return $Value
            }
        }

        default {
            return $Value
        }
    }
}

function Resolve-HubspotPropertyDisplayValue {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [Parameter(Mandatory = $false)]$Value,
        [Parameter(Mandatory = $false)]$Definition,
        [Parameter(Mandatory = $false)][string]$PipelineId,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($PropertyName -ieq 'dealstage') {
        $stageDefinition = Get-HubspotDealStageDefinition -PipelineId $PipelineId -StageId ([string]$Value) -AccessToken $AccessToken
        if ($stageDefinition -and -not [string]::IsNullOrWhiteSpace([string]$stageDefinition.Label)) {
            return [string]$stageDefinition.Label
        }
        return $Value
    }

    if ($PropertyName -ieq 'hubspot_owner_id') {
        return Get-HubspotOwnerDisplayValue -OwnerId ([string]$Value) -AccessToken $AccessToken
    }

    if ($PropertyName -match '^(associatedcompanyid|companyid|hs_company_id)$') {
        return Get-HubspotRelatedObjectDisplayValue -PropertyName $PropertyName -ObjectType 'companies' -ObjectId ([string]$Value -replace ';.*$') -QueryProperties @('name', 'domain') -DisplayPropertyCandidates @('name', 'domain') -RequiredReadScope 'crm.objects.companies.read' -AccessToken $AccessToken
    }

    if ($PropertyName -match '^(associatedcontactid|contactid|hs_contact_id)$') {
        return Get-HubspotRelatedObjectDisplayValue -PropertyName $PropertyName -ObjectType 'contacts' -ObjectId ([string]$Value -replace ';.*$') -QueryProperties @('firstname', 'lastname', 'email') -DisplayPropertyCandidates @('email') -RequiredReadScope 'crm.objects.contacts.read' -AccessToken $AccessToken
    }

    if ($PropertyName -match '^(associatedlineitemid|line_item_id|hs_line_item_id)$') {
        return Get-HubspotRelatedObjectDisplayValue -PropertyName $PropertyName -ObjectType 'line_items' -ObjectId ([string]$Value -replace ';.*$') -QueryProperties @('name', 'hs_name') -DisplayPropertyCandidates @('name', 'hs_name') -RequiredReadScope 'crm.objects.line_items.read' -AccessToken $AccessToken
    }

    if ($PropertyName -match '^(associatedproductid|productid|hs_product_id)$') {
        return Get-HubspotRelatedObjectDisplayValue -PropertyName $PropertyName -ObjectType 'products' -ObjectId ([string]$Value -replace ';.*$') -QueryProperties @('name', 'hs_name') -DisplayPropertyCandidates @('name', 'hs_name') -RequiredReadScope 'crm.objects.products.read' -AccessToken $AccessToken
    }

    if ($PropertyName -match '^(associatedquoteid|quoteid|hs_quote_id)$') {
        return Get-HubspotRelatedObjectDisplayValue -PropertyName $PropertyName -ObjectType 'quotes' -ObjectId ([string]$Value -replace ';.*$') -QueryProperties @('hs_title', 'hs_quote_number') -DisplayPropertyCandidates @('hs_title', 'hs_quote_number') -RequiredReadScope 'crm.objects.quotes.read' -AccessToken $AccessToken
    }

    if ($PropertyName -match '^(associatedticketid|ticketid|hs_ticket_id)$') {
        return Get-HubspotRelatedObjectDisplayValue -PropertyName $PropertyName -ObjectType 'tickets' -ObjectId ([string]$Value -replace ';.*$') -QueryProperties @('subject', 'content') -DisplayPropertyCandidates @('subject', 'content') -RequiredReadScope 'crm.objects.tickets.read' -AccessToken $AccessToken
    }

    if ($Definition -and [string]$Definition.type -eq 'enumeration' -and $Definition.options) {
        $optionLabelByValue = @{}
        foreach ($option in @($Definition.options)) {
            if ($option -and $option.value) {
                $optionLabelByValue[[string]$option.value] = [string]$option.label
            }
        }

        if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
            $labels = @()
            foreach ($entry in $Value) {
                $entryString = [string]$entry
                if ($optionLabelByValue.ContainsKey($entryString)) {
                    $labels += $optionLabelByValue[$entryString]
                }
                else {
                    $labels += $entry
                }
            }
            return $labels
        }

        $valueString = [string]$Value
        if ($optionLabelByValue.ContainsKey($valueString)) {
            return $optionLabelByValue[$valueString]
        }
    }

    return $Value
}

function Convert-HubspotRecordToObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $false)][string]$ObjectType = 'deals',
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    $propertyDefinitions = Get-HubspotPropertyDefinitions -ObjectType $ObjectType -AccessToken $AccessToken
    $convertedRecord = [ordered]@{}
    $pipelineId = $null
    $dealStageId = $null

    if ($Record.PSObject.Properties.Name -contains 'id') {
        $convertedRecord.Id = [string]$Record.id
    }

    if ($Record.PSObject.Properties.Name -contains 'archived') {
        $convertedRecord.Archived = [bool]$Record.archived
    }

    foreach ($metaPropertyName in @('createdAt', 'updatedAt', 'createdBy', 'updatedBy', 'propertiesWithHistory', 'associations')) {
        if ($Record.PSObject.Properties.Name -contains $metaPropertyName) {
            $convertedRecord[$metaPropertyName] = Convert-HubspotPropertyValue -Value $Record.$metaPropertyName
        }
    }

    if ($Record.PSObject.Properties.Name -contains 'properties' -and $Record.properties) {
        foreach ($property in $Record.properties.PSObject.Properties) {
            $definition = $null
            if ($propertyDefinitions.ContainsKey($property.Name)) {
                $definition = $propertyDefinitions[$property.Name]
            }

            $convertedValue = Convert-HubspotPropertyValue -Value $property.Value -Definition $definition
            $convertedRecord[$property.Name] = $convertedValue

            if ($property.Name -ieq 'pipeline') {
                $pipelineId = [string]$convertedValue
            }

            if ($property.Name -ieq 'dealstage') {
                $dealStageId = [string]$convertedValue
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($dealStageId)) {
        $dealStageDefinition = Get-HubspotDealStageDefinition -PipelineId $pipelineId -StageId $dealStageId -AccessToken $AccessToken
        if ($dealStageDefinition) {
            $convertedRecord.dealstageLabel = $dealStageDefinition.Label

            if ([string]::IsNullOrWhiteSpace($convertedRecord.pipelineLabel) -and -not [string]::IsNullOrWhiteSpace($dealStageDefinition.PipelineLabel)) {
                $convertedRecord.pipelineLabel = $dealStageDefinition.PipelineLabel
            }
        }
    }

    return [pscustomobject]$convertedRecord
}

<#
.SYNOPSIS
Stores the HubSpot bearer token securely in Windows Credential Manager.

.DESCRIPTION
Saves the token as a secure credential entry for local reuse by module commands.
You can pass a secure string or plain string token.

.EXAMPLE
$secureToken = Read-Host 'HubSpot Token' -AsSecureString
Set-HubspotAccessToken -SecureToken $secureToken

.EXAMPLE
Set-HubspotAccessToken -SecureToken (Read-Host 'HubSpot Token' -AsSecureString) -ServiceUserName 'Hubspot_Deals' -Target 'Hubspot_Deals_BearerToken'

.EXAMPLE
# Troubleshooting: Bei INVALID_AUTHENTICATION Token rotieren und neu speichern.
$secureToken = Read-Host 'Neuer HubSpot Token' -AsSecureString
Set-HubspotAccessToken -SecureToken $secureToken
#>
function Set-HubspotAccessToken {
    [CmdletBinding(DefaultParameterSetName = 'Secure')]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'Secure')]
        [Security.SecureString]$SecureToken,

        [Parameter(Mandatory = $true, ParameterSetName = 'Plain')]
        [string]$Token,

        [Parameter(Mandatory = $false)]
        [string]$ServiceUserName = $script:HubspotSettings.ServiceUserName,

        [Parameter(Mandatory = $false)]
        [string]$Target = $script:HubspotSettings.TokenTarget
    )

    $secureValue = $SecureToken
    if ($PSCmdlet.ParameterSetName -eq 'Plain') {
        $secureValue = ConvertTo-SecureString -String $Token -AsPlainText -Force
    }

    $credential = [PSCredential]::new($ServiceUserName, $secureValue)
    Set-Credential -Target $Target -Credential $credential -Type Generic -Persistence Enterprise -Description 'HubSpot Deals API token' > $null
}

function Get-HubspotBearerToken {
    param (
        [Parameter(Mandatory = $false)]
        [string]$AccessToken,

        [Parameter(Mandatory = $false)]
        [switch]$ForcePrompt
    )

    if (-not [string]::IsNullOrWhiteSpace($AccessToken)) {
        return $AccessToken
    }

    Initialize-AutomationEnvironment

    if ($script:env_runbook) {
        $automationCredentialCommand = Get-Command -Name Get-AutomationPSCredential -ErrorAction SilentlyContinue
        if (-not $automationCredentialCommand) {
            throw 'Get-AutomationPSCredential ist in dieser Runbook-Umgebung nicht verfuegbar.'
        }

        try {
            $credential = Get-AutomationPSCredential -Name $script:HubspotSettings.ServiceUserName -ErrorAction Stop
        }
        catch {
            throw "AutomationPSCredential mit dem Namen $($script:HubspotSettings.ServiceUserName) konnte nicht gelesen werden."
        }

        if (-not $credential) {
            throw "AutomationPSCredential mit dem Namen $($script:HubspotSettings.ServiceUserName) wurde nicht gefunden."
        }

        return Convert-SecureStringToPlainText -SecureString $credential.Password
    }

    if (-not $ForcePrompt.IsPresent) {
        $storedCredential = Get-StoredCredentialSafe -Target $script:HubspotSettings.TokenTarget
        if ($storedCredential) {
            return Convert-SecureStringToPlainText -SecureString $storedCredential.Password
        }
    }

    if ($script:env_runbook) {
        throw "Kein HubSpot Token gefunden. Hinterlegen Sie ihn mit Set-HubspotAccessToken oder uebergeben Sie -AccessToken direkt."
    }

    if ($ForcePrompt.IsPresent) {
        Write-Warning 'Der gespeicherte HubSpot Access-Token ist ungueltig oder abgelaufen. Bitte geben Sie den aktuellen Token erneut ein. Er wird danach gespeichert.'
    }
    else {
        Write-Host 'HubSpot Access-Token wird benoetigt' -ForegroundColor Yellow
    }

    $credential = Microsoft.PowerShell.Security\Get-Credential -UserName $script:HubspotSettings.ServiceUserName -Message 'Geben Sie den HubSpot Access-Token ein'
    if (-not $credential) {
        throw 'HubSpot Access-Token wurde nicht eingegeben.'
    }

    Set-HubspotAccessToken -SecureToken $credential.Password -ServiceUserName $script:HubspotSettings.ServiceUserName -Target $script:HubspotSettings.TokenTarget
    return Convert-SecureStringToPlainText -SecureString $credential.Password
}

<#
.SYNOPSIS
Executes a HubSpot Deals API request via the module REST wrapper.

.DESCRIPTION
Builds the deals endpoint path automatically and sends authenticated HTTP requests.
This command is useful for advanced/custom requests against deals resources.

.EXAMPLE
Invoke-HubspotRest

.EXAMPLE
Invoke-HubspotRest -Path '12345678901'

.EXAMPLE
Invoke-HubspotRest -Path 'search' -Method 'POST' -Body @{
    filterGroups = @(@{ filters = @(@{ propertyName = 'dealname'; operator = 'CONTAINS_TOKEN'; value = 'Test' }) })
    limit = 20
}

.EXAMPLE
Invoke-HubspotRest -Path '12345678901' -Method 'PATCH' -Body @{ properties = @{ dealname = 'Updated via Invoke-HubspotRest' } }

.EXAMPLE
# Troubleshooting: HubSpot-Fehlerdetails fuer Status/Scope-Diagnose sichtbar machen.
try {
    Invoke-HubspotRest -Path 'search' -Method 'POST' -Body @{ filterGroups = @(@{ filters = @(@{ propertyName = 'dealname'; operator = 'CONTAINS_TOKEN'; value = 'Test' }) }); limit = 1 } -ErrorAction Stop
}
catch {
    $PSItem | Format-List * -Force
}
#>
function Invoke-HubspotRest {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)][string]$Path,
        [Parameter(Mandatory = $false)][string]$Method = 'GET',
        [Parameter(Mandatory = $false)][hashtable]$Query,
        [Parameter(Mandatory = $false)]$Body,
        [Parameter(Mandatory = $false)][string]$AccessToken,
        [Parameter(Mandatory = $false)][int]$RetryCount = $script:HubspotSettings.DefaultRetryCount,
        [Parameter(Mandatory = $false)][double]$ThrottleDelaySeconds = $script:HubspotSettings.DefaultThrottleDelaySeconds
    )

    $resourcePath = (Get-HubspotDealsPath).Trim('/ ')
    if ($Path) {
        $resourcePath = "$resourcePath/$($Path.TrimStart('/'))"
    }

    return Invoke-HubspotRequest -ResourcePath $resourcePath -Method $Method -Query $Query -Body $Body -AccessToken $AccessToken -RetryCount $RetryCount -ThrottleDelaySeconds $ThrottleDelaySeconds
}

<#
.SYNOPSIS
Gets a single deal by ID.

.DESCRIPTION
Reads one deal from HubSpot. By default HubSpot standard attributes are returned.
Use -AllProperties to explicitly request all known deal properties.

.EXAMPLE
Get-HubspotDeal -Id '12345678901'

.EXAMPLE
Get-HubspotDeal -Id '12345678901' -AllProperties

.EXAMPLE
Get-HubspotDeal -Id '12345678901' -Properties @('dealname','pipeline','dealstage','hubspot_owner_id')

.EXAMPLE
Get-HubspotDeal -Id '12345678901' -PropertiesWithHistory @('dealstage') -Associations @('companies','contacts')

.EXAMPLE
Get-HubspotDeal -Id '12345678901' -Archived

.EXAMPLE
# Deal ueber ein eindeutiges Property statt ueber die Deal-ID lesen (z. B. ZEP-Angebotsnummer).
Get-HubspotDeal -Id '18380' -IdProperty 'zep_angebots_id' -Properties @('dealname','zep_angebots_id')

.EXAMPLE
# Troubleshooting: Bei 404 den Datensatz explizit als archiviert abrufen.
$id = '12345678901'
try {
    Get-HubspotDeal -Id $id -ErrorAction Stop
}
catch {
    Get-HubspotDeal -Id $id -Archived
}
#>
function Get-HubspotDeal {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $false)][string]$IdProperty,
        [Parameter(Mandatory = $false)][string[]]$Properties,
        [Parameter(Mandatory = $false)][switch]$AllProperties,
        [Parameter(Mandatory = $false)][string[]]$PropertiesWithHistory,
        [Parameter(Mandatory = $false)][string[]]$Associations,
        [Parameter(Mandatory = $false)][switch]$Archived,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    $query = @{}
    if ($Properties) {
        $query.properties = $Properties
    }

    if ($AllProperties.IsPresent -and -not $Properties) {
        $propertyDefinitions = Get-HubspotPropertyDefinitions -ObjectType 'deals' -AccessToken $AccessToken
        if ($propertyDefinitions -and $propertyDefinitions.Keys.Count -gt 0) {
            $query.properties = @($propertyDefinitions.Keys | Sort-Object)
        }
    }

    if ($PropertiesWithHistory) {
        $query.propertiesWithHistory = $PropertiesWithHistory
    }

    if ($Associations) {
        $query.associations = $Associations
    }

    if ($Archived.IsPresent) {
        $query.archived = 'true'
    }

    if (-not [string]::IsNullOrWhiteSpace($IdProperty)) {
        $query.idProperty = $IdProperty
    }

    return Convert-HubspotRecordToObject -Record (Invoke-HubspotRest -Path $Id -Query $query -AccessToken $AccessToken) -AccessToken $AccessToken
}

<#
.SYNOPSIS
Returns GUI-like detailed property rows for one deal.

.DESCRIPTION
Builds a property-focused output with labels, groups, types, raw values,
and display values (for example translated stage labels and owner names).

.EXAMPLE
Get-HubspotDealDetails -Id '12345678901'

.EXAMPLE
Get-HubspotDealDetails -Id '12345678901' -Properties @('dealname','pipeline','dealstage','hubspot_owner_id')

.EXAMPLE
Get-HubspotDealDetails -Id '12345678901' -Properties @('nda_safe') -IncludeEmpty

.EXAMPLE
Get-HubspotDealDetails -Id '12345678901' -Archived

.EXAMPLE
# Troubleshooting: Wenn DisplayValue nur ID bleibt, Scope-Hinweise pruefen.
$rows = Get-HubspotDealDetails -Id '12345678901' -Properties @('hubspot_owner_id')
$rows | Select-Object PropertyName, Value, DisplayValue
#>
function Get-HubspotDealDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $false)][string[]]$Properties,
        [Parameter(Mandatory = $false)][switch]$IncludeEmpty,
        [Parameter(Mandatory = $false)][switch]$Archived,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    $propertyDefinitions = Get-HubspotPropertyDefinitions -ObjectType 'deals' -AccessToken $AccessToken

    $propertyNames = @()
    $explicitPropertiesRequested = $false
    if ($Properties -and $Properties.Count -gt 0) {
        $explicitPropertiesRequested = $true
        $propertyNames = @($Properties | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    else {
        $propertyNames = @($propertyDefinitions.Keys | Sort-Object)
    }

    $query = @{}
    if ($propertyNames.Count -gt 0) {
        $query.properties = $propertyNames
    }

    if ($Archived.IsPresent) {
        $query.archived = 'true'
    }

    $record = $null
    try {
        $record = Invoke-HubspotRest -Path $Id -Query $query -AccessToken $AccessToken
    }
    catch {
        $statusCode = $null
        if ($PSItem.Exception.Response -and $PSItem.Exception.Response.StatusCode) {
            $statusCode = [int]$PSItem.Exception.Response.StatusCode
        }

        if (-not $Archived.IsPresent -and $statusCode -eq 404) {
            $query.archived = 'true'
            $record = Invoke-HubspotRest -Path $Id -Query $query -AccessToken $AccessToken
        }
        else {
            throw
        }
    }

    $converted = Convert-HubspotRecordToObject -Record $record -AccessToken $AccessToken

    $pipelineId = $null
    if ($record -and $record.properties -and ($record.properties.PSObject.Properties.Name -contains 'pipeline')) {
        $pipelineId = [string]$record.properties.pipeline
    }

    $details = @()
    foreach ($propertyName in $propertyNames) {
        $definition = $null
        if ($propertyDefinitions.ContainsKey($propertyName)) {
            $definition = $propertyDefinitions[$propertyName]
        }

        $value = $null
        if ($converted.PSObject.Properties.Name -contains $propertyName) {
            $value = $converted.$propertyName
        }

        if (-not $IncludeEmpty.IsPresent -and -not $explicitPropertiesRequested) {
            $isEmptyString = ($value -is [string] -and [string]::IsNullOrWhiteSpace($value))
            if ($null -eq $value -or $isEmptyString) {
                continue
            }
        }

        $displayValue = Resolve-HubspotPropertyDisplayValue -PropertyName $propertyName -Value $value -Definition $definition -PipelineId $pipelineId -AccessToken $AccessToken

        $details += [pscustomobject]@{
            Id = $Id
            PropertyName = $propertyName
            Label = if ($definition -and $definition.label) { [string]$definition.label } else { $propertyName }
            GroupName = if ($definition -and $definition.groupName) { [string]$definition.groupName } else { $null }
            Type = if ($definition -and $definition.type) { [string]$definition.type } else { $null }
            FieldType = if ($definition -and $definition.fieldType) { [string]$definition.fieldType } else { $null }
            Value = $value
            DisplayValue = $displayValue
        }
    }

    return $details
}

<#
.SYNOPSIS
Gets one page or all pages of deals.

.DESCRIPTION
Lists deals with paging support. Use -All to follow paging cursors.
Use -AllProperties to request all known deal properties.

.EXAMPLE
Get-HubspotDeals

.EXAMPLE
Get-HubspotDeals -Limit 50

.EXAMPLE
Get-HubspotDeals -All

.EXAMPLE
Get-HubspotDeals -AllProperties -Limit 100

.EXAMPLE
Get-HubspotDeals -Properties @('dealname','dealstage','pipeline','hubspot_owner_id') -All

.EXAMPLE
Get-HubspotDeals -Archived -All

.EXAMPLE
Get-HubspotDeals -After 'NTI4MjQ2MjA3Mjg='

.EXAMPLE
# Troubleshooting: Bei Rate-Limits kleinere Seiten laden und alle Seiten iterieren.
Get-HubspotDeals -Limit 20 -All
#>
function Get-HubspotDeals {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)][int]$Limit = $script:HubspotSettings.DefaultPageSize,
        [Parameter(Mandatory = $false)][string]$After,
        [Parameter(Mandatory = $false)][string[]]$Properties,
        [Parameter(Mandatory = $false)][switch]$AllProperties,
        [Parameter(Mandatory = $false)][string[]]$PropertiesWithHistory,
        [Parameter(Mandatory = $false)][string[]]$Associations,
        [Parameter(Mandatory = $false)][switch]$Archived,
        [Parameter(Mandatory = $false)][switch]$All,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    $safeLimit = [Math]::Max(1, [Math]::Min(200, $Limit))
    $allResults = @()
    $cursor = $After

    do {
        $query = @{
            limit = $safeLimit
        }

        if (-not [string]::IsNullOrWhiteSpace($cursor)) {
            $query.after = $cursor
        }

        if ($Properties) {
            $query.properties = $Properties
        }
        elseif ($AllProperties.IsPresent) {
            $propertyDefinitions = Get-HubspotPropertyDefinitions -ObjectType 'deals' -AccessToken $AccessToken
            if ($propertyDefinitions -and $propertyDefinitions.Keys.Count -gt 0) {
                $query.properties = @($propertyDefinitions.Keys | Sort-Object)
            }
        }

        if ($PropertiesWithHistory) {
            $query.propertiesWithHistory = $PropertiesWithHistory
        }

        if ($Associations) {
            $query.associations = $Associations
        }

        if ($Archived.IsPresent) {
            $query.archived = 'true'
        }

        $response = Invoke-HubspotRest -Query $query -AccessToken $AccessToken
        $convertedResults = @($response.results | ForEach-Object { Convert-HubspotRecordToObject -Record $_ -AccessToken $AccessToken })
        $allResults += $convertedResults

        if (-not $All.IsPresent) {
            return $convertedResults
        }

        $cursor = $null
        if ($response.paging -and $response.paging.next -and $response.paging.next.after) {
            $cursor = [string]$response.paging.next.after
        }
    }
    while (-not [string]::IsNullOrWhiteSpace($cursor))

    return $allResults
}

<#
.SYNOPSIS
Searches deals by property value.

.DESCRIPTION
Runs a deals search against the selected property. Default operator is CONTAINS_TOKEN;
use -Operator EQ for exact matches (e.g. unique IDs), HAS_PROPERTY to find deals where a property is set,
or IN / NOT_IN with -Values for several values. -SortProperty / -SortDescending control the order.
Supports paging via -All and optional selection of returned properties.

.EXAMPLE
Search-HubspotDeals -PropertyName 'dealname' -SearchTerm 'Mustermann'

.EXAMPLE
Search-HubspotDeals -PropertyName 'dealname' -SearchTerm 'Test' -Limit 20

.EXAMPLE
Search-HubspotDeals -PropertyName 'dealname' -SearchTerm 'Test' -Properties @('dealname','pipeline','dealstage')

.EXAMPLE
Search-HubspotDeals -PropertyName 'dealname' -SearchTerm 'Test' -All

.EXAMPLE
Search-HubspotDeals -PropertyName 'dealname' -SearchTerm 'Test' -Archived

.EXAMPLE
Search-HubspotDeals -PropertyName 'zep_angebots_id' -Operator 'EQ' -SearchTerm '18380'

.EXAMPLE
Search-HubspotDeals -PropertyName 'zep_angebots_id' -Operator 'HAS_PROPERTY' -All

.EXAMPLE
# Alle Deals, deren ZEP-Status neu (10) oder in Arbeit (20) ist.
Search-HubspotDeals -PropertyName 'zep_status' -Operator 'IN' -Values '10', '20' -Properties @('zep_angebots_id', 'zep_status') -All

.EXAMPLE
# Hoechste ZEP-Angebotsnummer in HubSpot.
Search-HubspotDeals -PropertyName 'zep_angebots_id' -Operator 'HAS_PROPERTY' -SortProperty 'zep_angebots_id' -SortDescending -Limit 1

.EXAMPLE
# Troubleshooting: Ungueltige Such-Properties schnell validieren.
try {
    Search-HubspotDeals -PropertyName 'Archived' -SearchTerm 'true' -ErrorAction Stop
}
catch {
    $PSItem.Exception.Message
}
#>
function Search-HubspotDeals {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [Parameter(Mandatory = $false)][string]$SearchTerm,
        [Parameter(Mandatory = $false)][ValidateSet('CONTAINS_TOKEN', 'NOT_CONTAINS_TOKEN', 'EQ', 'NEQ', 'LT', 'LTE', 'GT', 'GTE', 'HAS_PROPERTY', 'NOT_HAS_PROPERTY', 'IN', 'NOT_IN')][string]$Operator = 'CONTAINS_TOKEN',
        [Parameter(Mandatory = $false)][string[]]$Values,
        [Parameter(Mandatory = $false)][string]$SortProperty,
        [Parameter(Mandatory = $false)][switch]$SortDescending,
        [Parameter(Mandatory = $false)][int]$Limit = $script:HubspotSettings.DefaultPageSize,
        [Parameter(Mandatory = $false)][string]$After,
        [Parameter(Mandatory = $false)][string[]]$Properties,
        [Parameter(Mandatory = $false)][switch]$Archived,
        [Parameter(Mandatory = $false)][switch]$All,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    if ($PropertyName -ieq 'Archived') {
        throw 'Archived ist kein durchsuchbares Deal-Property. Verwenden Sie stattdessen -Archived mit einem echten Property-Namen oder Get-HubspotDeals -Archived.'
    }

    $operatorNeedsValues = $Operator -in @('IN', 'NOT_IN')
    $operatorNeedsValue = $Operator -notin @('HAS_PROPERTY', 'NOT_HAS_PROPERTY', 'IN', 'NOT_IN')
    if ($operatorNeedsValue -and [string]::IsNullOrWhiteSpace($SearchTerm)) {
        throw "Der Operator '$Operator' benoetigt einen -SearchTerm."
    }
    if ($operatorNeedsValues -and -not $Values) {
        throw "Der Operator '$Operator' benoetigt -Values."
    }

    $filter = @{
        propertyName = $PropertyName
        operator = $Operator
    }
    if ($operatorNeedsValue) {
        $filter.value = $SearchTerm
    }
    if ($operatorNeedsValues) {
        $filter.values = @($Values)
    }

    $safeLimit = [Math]::Max(1, [Math]::Min(200, $Limit))
    $allResults = @()
    $cursor = $After

    do {
        $body = @{
            filterGroups = @(
                @{
                    filters = @($filter)
                }
            )
            limit = $safeLimit
        }

        if (-not [string]::IsNullOrWhiteSpace($cursor)) {
            $body.after = $cursor
        }

        if ($Properties) {
            $body.properties = $Properties
        }

        if ($Archived.IsPresent) {
            $body.archived = $true
        }

        if (-not [string]::IsNullOrWhiteSpace($SortProperty)) {
            $body.sorts = @(@{ propertyName = $SortProperty; direction = $(if ($SortDescending.IsPresent) { 'DESCENDING' } else { 'ASCENDING' }) })
        }

        $response = Invoke-HubspotRest -Path 'search' -Method 'POST' -Body $body -AccessToken $AccessToken
        $convertedResults = @($response.results | ForEach-Object { Convert-HubspotRecordToObject -Record $_ -AccessToken $AccessToken })
        $allResults += $convertedResults

        if (-not $All.IsPresent) {
            return $convertedResults
        }

        $cursor = $null
        if ($response.paging -and $response.paging.next -and $response.paging.next.after) {
            $cursor = [string]$response.paging.next.after
        }
    }
    while (-not [string]::IsNullOrWhiteSpace($cursor))

    return $allResults
}

<#
.SYNOPSIS
Creates a new deal.

.DESCRIPTION
Creates a deal from the provided property hashtable. Optional associations
can be provided when needed.

.EXAMPLE
New-HubspotDeal -Properties @{
    dealname = 'Neuer Deal'
    pipeline = 'default'
    dealstage = 'appointmentscheduled'
}

.EXAMPLE
New-HubspotDeal -Properties @{
    dealname = 'Neuer Deal mit Betrag'
    amount = '5000'
    closedate = '2026-12-31'
}

.EXAMPLE
New-HubspotDeal -Properties @{
    dealname = 'Deal mit Eigentuemer'
    hubspot_owner_id = '12345678'
}

.EXAMPLE
# Deal direkt mit primaerer Firma anlegen.
New-HubspotDeal -Properties @{ dealname = 'Deal mit Firma'; pipeline = 'default'; dealstage = '1110595160' } -PrimaryCompanyId '10000000001'

.EXAMPLE
# Troubleshooting: Stage-Mapping pruefen und bei Fehlern mit interner Stage-ID anlegen.
New-HubspotDeal -Properties @{
    dealname = 'Fallback mit Stage-ID'
    pipeline = 'default'
    dealstage = 'appointmentscheduled'
}
#>
function New-HubspotDeal {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][hashtable]$Properties,
        [Parameter(Mandatory = $false)]$Associations,
        [Parameter(Mandatory = $false)][string]$PrimaryCompanyId,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    $body = @{
        properties = Convert-HubspotDealPropertiesForWrite -Properties $Properties -AccessToken $AccessToken
    }

    $associationList = @()
    if ($PSBoundParameters.ContainsKey('Associations') -and $null -ne $Associations) {
        $associationList += @($Associations)
    }

    if (-not [string]::IsNullOrWhiteSpace($PrimaryCompanyId)) {
        $associationList += @{
            to = @{ id = $PrimaryCompanyId }
            types = @(
                @{ associationCategory = 'HUBSPOT_DEFINED'; associationTypeId = $script:HubspotDealToPrimaryCompanyTypeId }
            )
        }
    }

    if ($associationList.Count -gt 0) {
        $body.associations = $associationList
    }

    return Convert-HubspotRecordToObject -Record (Invoke-HubspotRest -Method 'POST' -Body $body -AccessToken $AccessToken) -AccessToken $AccessToken
}

<#
.SYNOPSIS
Updates an existing deal.

.DESCRIPTION
Applies partial updates to one deal by ID using a property hashtable.

.EXAMPLE
Set-HubspotDeal -Id '12345678901' -Properties @{ dealname = 'Neuer Dealname' }

.EXAMPLE
Set-HubspotDeal -Id '12345678901' -Properties @{ dealstage = 'closedlost' }

.EXAMPLE
Set-HubspotDeal -Id '12345678901' -Properties @{ amount = '3000'; closedate = '2026-10-01' }

.EXAMPLE
Set-HubspotDeal -Id '12345678901' -Properties @{ hubspot_owner_id = '12345678' }

.EXAMPLE
# Typisierte Werte werden passend zum Property-Typ formatiert (date, datetime, bool, Zahl).
Set-HubspotDeal -Id '18380' -IdProperty 'zep_angebots_id' -Properties @{
    abgabe_des_angebotes__kundenwunsch_ = (Get-Date).AddDays(7)
    zep_status_since = [datetime]::UtcNow
    nda___safe = $false
}

.EXAMPLE
# Troubleshooting: Stage stabil ueber interne ID statt Label setzen.
Set-HubspotDeal -Id '12345678901' -Properties @{ pipeline = 'default'; dealstage = 'closedlost' }
#>
function Set-HubspotDeal {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $false)][string]$IdProperty,
        [Parameter(Mandatory = $true)][hashtable]$Properties,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    $body = @{
        properties = Convert-HubspotDealPropertiesForWrite -Properties $Properties -AccessToken $AccessToken
    }

    $query = @{}
    if (-not [string]::IsNullOrWhiteSpace($IdProperty)) {
        $query.idProperty = $IdProperty
    }

    return Convert-HubspotRecordToObject -Record (Invoke-HubspotRest -Path $Id -Method 'PATCH' -Query $query -Body $body -AccessToken $AccessToken) -AccessToken $AccessToken
}

<#
.SYNOPSIS
Deletes a deal.

.DESCRIPTION
Deletes a deal by ID. Supports ShouldProcess and confirmation prompts.

.EXAMPLE
Remove-HubspotDeal -Id '12345678901'

.EXAMPLE
Remove-HubspotDeal -Id '12345678901' -Confirm:$false

.EXAMPLE
Remove-HubspotDeal -Id '12345678901' -WhatIf

.EXAMPLE
# Troubleshooting: Vor endgueltigem Loeschen zuerst Simulationslauf ausfuehren.
Remove-HubspotDeal -Id '12345678901' -WhatIf
Remove-HubspotDeal -Id '12345678901' -Confirm:$false
#>
function Remove-HubspotDeal {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    if ($PSCmdlet.ShouldProcess("Deal $Id", 'Delete')) {
        Invoke-HubspotRest -Path $Id -Method 'DELETE' -AccessToken $AccessToken > $null
    }
}

<#
.SYNOPSIS
Gets a HubSpot owner by email or owner ID.

.DESCRIPTION
Looks up an owner (user who can own CRM records) and returns Id, Email, FirstName, LastName,
UserId and Archived. Results are cached per session; use -Refresh to bypass the cache.
Returns $null when no owner matches.

.EXAMPLE
Get-HubspotOwner -Email 'max.mustermann@contoso.com'

.EXAMPLE
(Get-HubspotOwner -Email 'max.mustermann@contoso.com').Id

.EXAMPLE
Get-HubspotOwner -Id '12345678'

.EXAMPLE
# Troubleshooting: Cache umgehen, wenn ein Owner gerade erst angelegt wurde.
Get-HubspotOwner -Email 'neu@contoso.com' -Refresh
#>
function Get-HubspotOwner {
    [CmdletBinding(DefaultParameterSetName = 'ByEmail')]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'ByEmail')][string]$Email,
        [Parameter(Mandatory = $true, ParameterSetName = 'ById')][string]$Id,
        [Parameter(Mandatory = $false)][switch]$Refresh,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    if (-not $script:HubspotOwnerCache) {
        $script:HubspotOwnerCache = @{}
    }

    $cacheKey = if ($PSCmdlet.ParameterSetName -eq 'ByEmail') { "email:$($Email.Trim().ToLowerInvariant())" } else { "id:$Id" }
    if (-not $Refresh.IsPresent -and $script:HubspotOwnerCache.ContainsKey($cacheKey)) {
        return $script:HubspotOwnerCache[$cacheKey]
    }

    $owner = $null
    if ($PSCmdlet.ParameterSetName -eq 'ByEmail') {
        $response = Invoke-HubspotRequest -ResourcePath (Get-HubspotCrmPath -Api 'owners') -Query @{ email = $Email.Trim() } -AccessToken $AccessToken
        $owner = @($response.results) | Select-Object -First 1
    }
    else {
        try {
            $owner = Invoke-HubspotRequest -ResourcePath (Get-HubspotCrmPath -Api 'owners' -Resource $Id) -AccessToken $AccessToken
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
    }

    $result = $null
    if ($owner) {
        $result = [pscustomobject]@{
            Id = [string]$owner.id
            Email = [string]$owner.email
            FirstName = [string]$owner.firstName
            LastName = [string]$owner.lastName
            UserId = [string]$owner.userId
            Archived = [bool]$owner.archived
        }
    }

    $script:HubspotOwnerCache[$cacheKey] = $result
    return $result
}

<#
.SYNOPSIS
Lists the companies associated with a deal.

.DESCRIPTION
Returns one object per associated company with CompanyId, IsPrimary and the association labels.

.EXAMPLE
Get-HubspotDealCompanies -DealId '12345678903'

.EXAMPLE
Get-HubspotDealCompanies -DealId '12345678903' | Where-Object IsPrimary
#>
function Get-HubspotDealCompanies {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$DealId,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    $response = Invoke-HubspotRest -Path "$DealId/associations/companies" -AccessToken $AccessToken
    foreach ($association in @($response.results)) {
        if (-not $association) {
            continue
        }

        $typeIds = @($association.associationTypes | ForEach-Object { [int]$_.typeId })
        [pscustomobject]@{
            CompanyId = [string]$association.toObjectId
            IsPrimary = $typeIds -contains $script:HubspotDealToPrimaryCompanyTypeId
            Labels = @($association.associationTypes | Where-Object { $_.label } | ForEach-Object { [string]$_.label })
            TypeIds = $typeIds
        }
    }
}

<#
.SYNOPSIS
Sets the primary company of a deal.

.DESCRIPTION
Associates the company with the deal as primary company. HubSpot allows only one primary
company per deal; a previous primary company keeps a plain association unless
-RemovePreviousPrimary is used, which removes all associations to it.

.EXAMPLE
Set-HubspotDealPrimaryCompany -DealId '12345678903' -CompanyId '10000000001'

.EXAMPLE
# Kundenwechsel: neue primaere Firma setzen und die alte Firma vom Deal loesen.
Set-HubspotDealPrimaryCompany -DealId '12345678903' -CompanyId '10000000002' -RemovePreviousPrimary

.EXAMPLE
Set-HubspotDealPrimaryCompany -DealId '12345678903' -CompanyId '10000000002' -RemovePreviousPrimary -WhatIf
#>
function Set-HubspotDealPrimaryCompany {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)][string]$DealId,
        [Parameter(Mandatory = $true)][string]$CompanyId,
        [Parameter(Mandatory = $false)][switch]$RemovePreviousPrimary,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    $previousPrimary = @()
    if ($RemovePreviousPrimary.IsPresent) {
        $previousPrimary = @(Get-HubspotDealCompanies -DealId $DealId -AccessToken $AccessToken | Where-Object { $_.IsPrimary -and $_.CompanyId -ne $CompanyId })
    }

    if ($PSCmdlet.ShouldProcess("Deal $DealId", "Primaere Firma $CompanyId setzen")) {
        $body = @(
            @{ associationCategory = 'HUBSPOT_DEFINED'; associationTypeId = $script:HubspotDealToPrimaryCompanyTypeId },
            @{ associationCategory = 'HUBSPOT_DEFINED'; associationTypeId = $script:HubspotDealToCompanyTypeId }
        )
        $null = Invoke-HubspotRest -Path "$DealId/associations/companies/$CompanyId" -Method 'PUT' -Body $body -AccessToken $AccessToken
    }

    foreach ($company in $previousPrimary) {
        if ($PSCmdlet.ShouldProcess("Deal $DealId", "Verknuepfung zu bisheriger Firma $($company.CompanyId) entfernen")) {
            $null = Invoke-HubspotRest -Path "$DealId/associations/companies/$($company.CompanyId)" -Method 'DELETE' -AccessToken $AccessToken
        }
    }
}

<#
.SYNOPSIS
Gets a HubSpot company by ID.

.DESCRIPTION
Returns the company record (flattened like deals) or $null when the company does not exist
(also not archived). Useful to check whether an ID can be associated before creating a deal.

.EXAMPLE
Get-HubspotCompany -Id '10000000001'

.EXAMPLE
if (-not (Get-HubspotCompany -Id '10000000002')) { 'Firma fehlt in HubSpot' }

.EXAMPLE
Get-HubspotCompany -Id '10000000001' -Properties @('name', 'domain', 'zep')
#>
function Get-HubspotCompany {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $false)][string[]]$Properties = @('name', 'domain'),
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    try {
        $record = Invoke-HubspotRequest -ResourcePath (Get-HubspotCrmPath -Api 'objects' -Resource "companies/$Id") -Query @{ properties = $Properties } -AccessToken $AccessToken -RetryCount 0
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }

        if ($statusCode -eq 404) {
            return $null
        }

        throw
    }

    $company = [ordered]@{ Id = [string]$record.id }
    if ($record.properties) {
        foreach ($property in $record.properties.PSObject.Properties) {
            $company[$property.Name] = $property.Value
        }
    }

    return [pscustomobject]$company
}

<#
.SYNOPSIS
Gets the definition of one CRM property.

.DESCRIPTION
Returns the HubSpot property definition (name, label, type, fieldType, options, ...)
or $null when the property does not exist.

.EXAMPLE
Get-HubspotProperty -Name 'zep_status'

.EXAMPLE
(Get-HubspotProperty -Name 'zep_status').options | Select-Object value, label
#>
function Get-HubspotProperty {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $false)][string]$ObjectType = 'deals',
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    try {
        return Invoke-HubspotRequest -ResourcePath (Get-HubspotCrmPath -Api 'properties' -Resource "$ObjectType/$Name") -AccessToken $AccessToken -RetryCount 0
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }

        if ($statusCode -eq 404) {
            return $null
        }

        throw
    }
}

<#
.SYNOPSIS
Creates a CRM property.

.DESCRIPTION
Creates a property from a definition hashtable (name, label, type, fieldType, groupName and
optionally description, options, hasUniqueValue, formField, hidden). Requires the scope
crm.schemas.<object>.write.

.EXAMPLE
New-HubspotProperty -Definition @{
    name = 'zep_status'; label = 'ZEP-Status'; type = 'enumeration'; fieldType = 'select'; groupName = 'dealinformation'
    options = @(@{ value = '10'; label = 'neu'; displayOrder = 0 }, @{ value = '20'; label = 'in Arbeit'; displayOrder = 1 })
}

.EXAMPLE
New-HubspotProperty -Definition @{ name = 'demo'; label = 'Demo'; type = 'string'; fieldType = 'text'; groupName = 'dealinformation' } -WhatIf
#>
function New-HubspotProperty {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)][hashtable]$Definition,
        [Parameter(Mandatory = $false)][string]$ObjectType = 'deals',
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    foreach ($required in 'name', 'label', 'type', 'fieldType', 'groupName') {
        if (-not $Definition.ContainsKey($required)) {
            throw "Property-Definition ohne '$required'."
        }
    }

    if ($PSCmdlet.ShouldProcess("$ObjectType.$($Definition.name)", 'Property anlegen')) {
        $result = Invoke-HubspotRequest -ResourcePath (Get-HubspotCrmPath -Api 'properties' -Resource $ObjectType) -Method 'POST' -Body $Definition -AccessToken $AccessToken
        if ($script:HubspotPropertyDefinitionsByObjectType) {
            $script:HubspotPropertyDefinitionsByObjectType.Remove($ObjectType)
        }
        return $result
    }
}

<#
.SYNOPSIS
Adds an option to an enumeration property.

.DESCRIPTION
Appends a new option (value + label) to a select/checkbox/radio property, keeping all existing
options. When the value already exists nothing changes, unless -UpdateLabel is used: then the label
of the existing option is replaced (the value stays, so stored data remains valid).
Returns the updated property definition.

.EXAMPLE
Add-HubspotPropertyOption -Name 'zep_status' -Value '300' -Label 'Neuer ZEP-Status'

.EXAMPLE
Add-HubspotPropertyOption -Name 'zep_status' -Value '300' -Label 'Neuer ZEP-Status' -WhatIf

.EXAMPLE
# Anzeigenamen einer bestehenden Option aendern.
Add-HubspotPropertyOption -Name 'zep_status' -Value '30' -Label 'fertig' -UpdateLabel
#>
function Add-HubspotPropertyOption {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Value,
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $false)][switch]$UpdateLabel,
        [Parameter(Mandatory = $false)][string]$ObjectType = 'deals',
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    $property = Get-HubspotProperty -Name $Name -ObjectType $ObjectType -AccessToken $AccessToken
    if (-not $property) {
        throw "Property '$ObjectType.$Name' existiert nicht."
    }

    $options = @($property.options | Where-Object { $_ })
    $existing = $options | Where-Object { [string]$_.value -eq $Value } | Select-Object -First 1
    if ($existing -and (-not $UpdateLabel.IsPresent -or [string]$existing.label -ceq $Label)) {
        return $property
    }

    $newOptions = @($options | ForEach-Object {
        $optionLabel = if ([string]$_.value -eq $Value) { $Label } else { [string]$_.label }
        @{ value = [string]$_.value; label = $optionLabel; displayOrder = [int]$_.displayOrder; hidden = [bool]$_.hidden }
    })
    if (-not $existing) {
        $newOptions += @{ value = $Value; label = $Label; displayOrder = $newOptions.Count; hidden = $false }
    }

    $action = if ($existing) { "Bezeichnung von Option '$Value' auf '$Label' aendern" } else { "Option '$Value' ($Label) ergaenzen" }
    if ($PSCmdlet.ShouldProcess("$ObjectType.$Name", $action)) {
        $result = Invoke-HubspotRequest -ResourcePath (Get-HubspotCrmPath -Api 'properties' -Resource "$ObjectType/$Name") -Method 'PATCH' -Body @{ options = $newOptions } -AccessToken $AccessToken
        if ($script:HubspotPropertyDefinitionsByObjectType) {
            $script:HubspotPropertyDefinitionsByObjectType.Remove($ObjectType)
        }
        return $result
    }
}

<#
.SYNOPSIS
Removes all associations between a deal and a company.

.DESCRIPTION
Deletes every association (primary and labeled) between the deal and the given company.
Neither the deal nor the company is deleted.

.EXAMPLE
Remove-HubspotDealCompany -DealId '12345678904' -CompanyId '10000000001'

.EXAMPLE
# Primaere Firma vom Deal loesen.
Get-HubspotDealCompanies -DealId '12345678904' | Where-Object IsPrimary | ForEach-Object { Remove-HubspotDealCompany -DealId '12345678904' -CompanyId $_.CompanyId }
#>
function Remove-HubspotDealCompany {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)][string]$DealId,
        [Parameter(Mandatory = $true)][string]$CompanyId,
        [Parameter(Mandatory = $false)][string]$AccessToken
    )

    if ($PSCmdlet.ShouldProcess("Deal $DealId", "Verknuepfung zu Firma $CompanyId entfernen")) {
        $null = Invoke-HubspotRest -Path "$DealId/associations/companies/$CompanyId" -Method 'DELETE' -AccessToken $AccessToken
    }
}
