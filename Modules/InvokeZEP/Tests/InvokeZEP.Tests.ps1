$modulePath = Join-Path $PSScriptRoot '..\InvokeZEP.psd1'
Import-Module $modulePath -Force

$config = Get-ZEPConfiguration
if (-not $config -or [string]::IsNullOrWhiteSpace($config.BaseUri)) {
    throw 'Get-ZEPConfiguration returned an invalid configuration object.'
}

$commands = @('Get-ZEPConfiguration', 'Set-ZEPConfiguration', 'Invoke-ZEPRest', 'Get-ZEPOffer', 'Get-ZEPOffers', 'Get-ZEPOfferItems', 'Search-ZEPOffers', 'ConvertFrom-ZEPDateTime', 'ConvertTo-ZEPTypedOffer', 'Get-ZEPRecentOffers', 'Get-ZEPOffersById', 'Get-ZEPEmployee')
foreach ($commandName in $commands) {
    if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
        throw "Command '$commandName' was not exported from the module."
    }
}

$cacheChanged = & (Get-Command -Name Test-ZEPOffersCacheNeedsRefresh -ErrorAction Stop) -CacheObject ([pscustomobject]@{ generatedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); total = 10; items = @() }) -CurrentTotal 12
if ($cacheChanged -ne $true) {
    throw 'The cache refresh detection did not detect a changed total count.'
}

$cachePartial = & (Get-Command -Name Test-ZEPOffersCacheNeedsRefresh -ErrorAction Stop) -CacheObject ([pscustomobject]@{ generatedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); total = 10; items = @(@{ id = 1 }, @{ id = 2 }) }) -CurrentTotal 10
if ($cachePartial -ne $true) {
    throw 'The cache refresh detection did not detect a partial cache.'
}

$cacheComplete = & (Get-Command -Name Test-ZEPOffersCacheNeedsRefresh -ErrorAction Stop) -CacheObject ([pscustomobject]@{ generatedAtUtc = (Get-Date).ToUniversalTime().ToString('o'); total = 10; items = @(@{ id = 1 }, @{ id = 2 }, @{ id = 3 }, @{ id = 4 }, @{ id = 5 }, @{ id = 6 }, @{ id = 7 }, @{ id = 8 }, @{ id = 9 }, @{ id = 10 }) }) -CurrentTotal 10
if ($cacheComplete -ne $false) {
    throw 'The cache refresh detection incorrectly flagged a complete cache as stale.'
}

$cacheExpired = & (Get-Command -Name Test-ZEPOffersCacheNeedsRefresh -ErrorAction Stop) -CacheObject ([pscustomobject]@{ generatedAtUtc = (Get-Date).ToUniversalTime().AddMinutes(-61).ToString('o'); total = 10; items = @(@{ id = 1 }, @{ id = 2 }, @{ id = 3 }, @{ id = 4 }, @{ id = 5 }, @{ id = 6 }, @{ id = 7 }, @{ id = 8 }, @{ id = 9 }, @{ id = 10 }) }) -CurrentTotal 10 -CacheTtlMinutes 60
if ($cacheExpired -ne $true) {
    throw 'The cache refresh detection did not detect an expired cache based on TTL.'
}

$tempCachePath = Join-Path ([System.IO.Path]::GetTempPath()) ('InvokeZEP-search-test-' + [guid]::NewGuid().ToString('N') + '.json')
try {
    $cacheFixture = [pscustomobject]@{
        version = 1
        generatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        total = 3
        items = @(
            [pscustomobject]@{ id = 1; data = [pscustomobject]@{ id = 1; title = 'Test Angebot'; status = [pscustomobject]@{ name = 'Aktiv' } } },
            [pscustomobject]@{ id = 2; data = [pscustomobject]@{ id = 2; title = 'Anderes Angebot'; status = [pscustomobject]@{ name = 'Inaktiv' } } },
            [pscustomobject]@{ id = 3; data = [pscustomobject]@{ id = 3; title = 'Demo test case'; status = [pscustomobject]@{ name = 'Aktiv' } } }
        )
    }

    $cacheFixture | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $tempCachePath -Encoding UTF8

    $searchResults = & (Get-Command -Name Search-ZEPOffers -ErrorAction Stop) -PropertyName 'title' -SearchTerm 'test' -ReferenceSampleSize 1000 -MaxResults 5 -ValidateCache:$false -CachePath $tempCachePath
    if (@($searchResults).Count -ne 2) {
        throw 'Search-ZEPOffers did not return the expected amount of title matches.'
    }
}
finally {
    if (Test-Path -LiteralPath $tempCachePath) {
        Remove-Item -LiteralPath $tempCachePath -Force -ErrorAction SilentlyContinue
    }
}

$manifest = Test-ModuleManifest -Path (Join-Path $PSScriptRoot '..\InvokeZEP.psd1')
if ($manifest.Version.ToString() -ne '0.3.1') {
    throw "Unexpected module version $($manifest.Version)."
}

# --- ConvertFrom-ZEPDateTime: ZEP sends Europe/Berlin wall clock marked as 'Z'
$summer = ConvertFrom-ZEPDateTime -Value '2026-10-07T01:24:35.000000Z'
if ($summer.Kind -ne [DateTimeKind]::Utc -or $summer.ToString('yyyy-MM-ddTHH:mm:ss') -ne '2026-10-06T23:24:35') {
    throw "ConvertFrom-ZEPDateTime (Sommerzeit) lieferte $($summer.ToString('o'))."
}

$winter = ConvertFrom-ZEPDateTime -Value '2026-12-01T10:00:00Z'
if ($winter.ToString('yyyy-MM-ddTHH:mm:ss') -ne '2026-12-01T09:00:00') {
    throw "ConvertFrom-ZEPDateTime (Winterzeit) lieferte $($winter.ToString('o'))."
}

# PowerShell 7 delivers [datetime] Kind=Utc with the unchanged wall clock; some versions Kind=Local (shifted).
$fromUtcKind = ConvertFrom-ZEPDateTime -Value ([datetime]::new(2026, 10, 7, 1, 24, 35, [DateTimeKind]::Utc))
$fromLocalKind = ConvertFrom-ZEPDateTime -Value ([datetime]::new(2026, 10, 7, 1, 24, 35, [DateTimeKind]::Utc).ToLocalTime())
if ($fromUtcKind -ne $summer -or $fromLocalKind -ne $summer) {
    throw 'ConvertFrom-ZEPDateTime behandelt [datetime]-Eingaben nicht wie den Rohstring.'
}

$dateOnly = ConvertFrom-ZEPDateTime -Value '2026-12-31T00:00:00.000000Z' -AsDate
if ($dateOnly.ToString('yyyy-MM-dd') -ne '2026-12-31' -or $dateOnly.Kind -ne [DateTimeKind]::Utc) {
    throw "ConvertFrom-ZEPDateTime -AsDate lieferte $($dateOnly.ToString('o'))."
}

if ($null -ne (ConvertFrom-ZEPDateTime -Value $null) -or $null -ne (ConvertFrom-ZEPDateTime -Value '')) {
    throw 'ConvertFrom-ZEPDateTime sollte fuer leere Werte $null liefern.'
}

# --- ConvertTo-ZEPTypedOffer
$rawOffer = '{"id":18380,"customer_id":"10000000002","department_id":1,"title":"ZZ_TEST","status":{"id":10,"name":"neu"},"status_since":"2026-10-07T01:24:35.000000Z","valid_until":"2027-01-31T00:00:00.000000Z","order_date":null,"responsible_username":"m.mustermann","processor_username":"m.mustermann","default_validity":null}' | ConvertFrom-Json
$typedOffer = ConvertTo-ZEPTypedOffer -Offer $rawOffer
if ($typedOffer.id -isnot [int] -or $typedOffer.status.id -isnot [int] -or $typedOffer.customer_id -isnot [string]) {
    throw 'ConvertTo-ZEPTypedOffer setzt die ID-Typen nicht korrekt.'
}
if ($typedOffer.status_since -ne $summer -or $typedOffer.valid_until.ToString('yyyy-MM-dd') -ne '2027-01-31' -or $null -ne $typedOffer.order_date) {
    throw 'ConvertTo-ZEPTypedOffer wandelt die Datumswerte nicht korrekt.'
}
if ($typedOffer.title -ne 'ZZ_TEST' -or $typedOffer.processor_username -ne 'm.mustermann') {
    throw 'ConvertTo-ZEPTypedOffer veraendert unbeteiligte Felder.'
}

# --- Paging, id[] and query arrays with a mocked API
$zepModule = Get-Module InvokeZEP
& $zepModule {
    $script:MockCalls = [System.Collections.Generic.List[object]]::new()
    $script:MockOffers = @(1..250 | ForEach-Object { [pscustomobject]@{ id = $_ + 1000; title = "Angebot $_"; status = [pscustomobject]@{ id = 10; name = 'neu' }; status_since = '2026-10-07T01:24:35Z'; customer_id = '1' } })

    function script:Invoke-ZEPRest {
        param ([string]$Path, [string]$Method = 'GET', [hashtable]$Query, $Body, [int]$RetryCount, [double]$ThrottleDelaySeconds)
        $script:MockCalls.Add([pscustomobject]@{ Path = $Path; Query = $Query })
        $offers = $script:MockOffers
        if ($Query.ContainsKey('id[]')) {
            $wanted = @($Query['id[]'] | ForEach-Object { [int]$_ })
            $offers = @($offers | Where-Object { $wanted -contains $_.id })
        }
        $limit = [int]$Query.limit
        $page = [int]$Query.page
        $lastPage = [Math]::Max(1, [int][Math]::Ceiling($offers.Count / $limit))
        $data = @($offers | Select-Object -Skip (($page - 1) * $limit) -First $limit)
        return [pscustomobject]@{ data = $data; meta = [pscustomobject]@{ total = $offers.Count; last_page = $lastPage; current_page = $page } }
    }
}

$recent = @(Get-ZEPRecentOffers -AfterId 1245 -PageSize 100)
if (($recent.id -join ',') -ne '1246,1247,1248,1249,1250') {
    throw "Get-ZEPRecentOffers -AfterId lieferte: $($recent.id -join ',')"
}
$pagesRequested = @(& $zepModule { $script:MockCalls } | Where-Object { $_.Query.limit -eq 100 } | ForEach-Object { $_.Query.page })
if (($pagesRequested -join ',') -ne '3') {
    throw "Get-ZEPRecentOffers hat unnoetige Seiten gelesen: $($pagesRequested -join ',')"
}

& $zepModule { $script:MockCalls.Clear() }
$recentAcrossPages = @(Get-ZEPRecentOffers -AfterId 1140 -PageSize 100)
if ($recentAcrossPages.Count -ne 110 -or $recentAcrossPages[0].id -ne 1141 -or $recentAcrossPages[-1].id -ne 1250) {
    throw 'Get-ZEPRecentOffers liest ueber Seitengrenzen nicht korrekt.'
}

$lastFive = @(Get-ZEPRecentOffers -Last 5 -PageSize 100)
if (($lastFive.id -join ',') -ne '1246,1247,1248,1249,1250' -or $lastFive[0].status_since -ne $summer) {
    throw 'Get-ZEPRecentOffers -Last liefert falsche oder untypisierte Angebote.'
}

& $zepModule { $script:MockCalls.Clear() }
$byId = @(Get-ZEPOffersById -Id 1100, 1001, 1100, 99999)
if (($byId.id -join ',') -ne '1001,1100') {
    throw "Get-ZEPOffersById lieferte: $($byId.id -join ',')"
}
$idCall = & $zepModule { $script:MockCalls[0] }
if ((@($idCall.Query['id[]']) -join ',') -ne '1100,1001,99999') {
    throw 'Get-ZEPOffersById uebergibt die IDs nicht als Liste.'
}

$manyIds = @(Get-ZEPOffersById -Id (1001..1250))
if ($manyIds.Count -ne 250 -or @(& $zepModule { $script:MockCalls } | Where-Object { $_.Query.ContainsKey('id[]') }).Count -ne 4) {
    throw 'Get-ZEPOffersById teilt grosse ID-Listen nicht in Bloecke zu 100.'
}

Remove-Module InvokeZEP -Force
Import-Module $modulePath -Force

# --- Invoke-ZEPRest builds repeated keys for array values
& (Get-Module InvokeZEP) {
    function script:Get-ZEPBearerToken { return 'test-token' }
    function script:Invoke-RestMethod { param ([uri]$Uri, [string]$Method, [hashtable]$Headers, $ErrorAction) $script:CapturedUri = $Uri.AbsoluteUri; return [pscustomobject]@{ data = @() } }
    $null = Invoke-ZEPRest -Path 'offers' -Query @{ 'id[]' = @(1, 2) } -ThrottleDelaySeconds 0
    if ($script:CapturedUri -notmatch 'id%5b%5d=1&id%5b%5d=2') {
        throw "Invoke-ZEPRest baut Listenparameter falsch: $($script:CapturedUri)"
    }
}

Remove-Module InvokeZEP -Force
Import-Module $modulePath -Force

Write-Host 'InvokeZEP module import and export validation passed.' -ForegroundColor Green
