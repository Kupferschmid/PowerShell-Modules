$modulePath = Join-Path $PSScriptRoot '..\InvokeZEP.psd1'
Import-Module $modulePath -Force

$config = Get-ZEPConfiguration
if (-not $config -or [string]::IsNullOrWhiteSpace($config.BaseUri)) {
    throw 'Get-ZEPConfiguration returned an invalid configuration object.'
}

$commands = @('Get-ZEPConfiguration', 'Set-ZEPConfiguration', 'Invoke-ZEPRest', 'Get-ZEPOffer', 'Get-ZEPOffers', 'Get-ZEPOfferItems', 'Search-ZEPOffers', 'ConvertFrom-ZEPDateTime', 'ConvertTo-ZEPTypedOffer', 'Get-ZEPRecentOffers', 'Get-ZEPOffersByStatus', 'Get-ZEPOffersById', 'Get-ZEPEmployee', 'Get-ZEPCustomerContact', 'Get-ZEPOfferCategories', 'Get-ZEPProject')
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
if ($manifest.Version.ToString() -ne '0.6.1') {
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
        if ($Query.orderBy -eq 'status' -and $script:MockHonorsSort) {
            $offers = @($offers | Sort-Object { $_.status.id }, { $_.id })
        }
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

# --- Get-ZEPOffersByStatus: 1001-1250 with mixed status, sorted by status via orderBy
& $zepModule {
    $script:MockCalls.Clear()
    $statusCycle = @(10, 20, 30, 60, 200)
    foreach ($offer in $script:MockOffers) {
        $statusValue = $statusCycle[($offer.id - 1001) % 5]
        $offer.status = [pscustomobject]@{ id = $statusValue; name = "s$statusValue" }
    }
    $script:MockHonorsSort = $true
}

$byStatus = @(Get-ZEPOffersByStatus -StatusId 10, 20 -AfterId 1200 -PageSize 50)
$expectedIds = @(1201..1250 | Where-Object { @(0, 1) -contains (($_ - 1001) % 5) })
if (($byStatus.id -join ',') -ne ($expectedIds -join ',') -or $byStatus[0].status.id -isnot [int]) {
    throw "Get-ZEPOffersByStatus lieferte: $($byStatus.id -join ',')"
}
# 50 offers per status: page 1 = status 10, page 2 = status 20, page 3 = status 30 -> stop
$statusPages = @(& $zepModule { $script:MockCalls } | ForEach-Object { $_.Query.page })
if (($statusPages -join ',') -ne '1,2,3') {
    throw "Get-ZEPOffersByStatus hat falsche Seiten gelesen: $($statusPages -join ',')"
}

& $zepModule { $script:MockHonorsSort = $false }
$unsortedFailed = $false
try { $null = Get-ZEPOffersByStatus -StatusId 10, 20 -PageSize 50 }
catch { $unsortedFailed = $_.Exception.Message -match 'nicht nach Status sortiert' }
if (-not $unsortedFailed) {
    throw 'Get-ZEPOffersByStatus bricht bei unsortierter Liste nicht ab.'
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

# --- Get-ZEPCustomerContact: path, session cache, $null for unknown contacts
& (Get-Module InvokeZEP) {
    $script:ContactCalls = [System.Collections.Generic.List[string]]::new()
    function script:Invoke-ZEPRest {
        param ([string]$Path, [string]$Method = 'GET', [hashtable]$Query, $Body, [int]$RetryCount, [double]$ThrottleDelaySeconds)
        $script:ContactCalls.Add($Path)
        if ($Path -eq 'customers/10000000001/contacts/12345') {
            return [pscustomobject]@{ data = [pscustomobject]@{ id = 12345; first_name = 'Max'; last_name = 'Mustermann'; email = 'max.mustermann@contoso.com' } }
        }
        $ex = [System.Exception]::new('Not Found')
        Add-Member -InputObject $ex -MemberType NoteProperty -Name Response -Value ([pscustomobject]@{ StatusCode = 404 }) -Force
        throw $ex
    }
}

$zepContact = Get-ZEPCustomerContact -CustomerId '10000000001' -ContactId '12345'
$null = Get-ZEPCustomerContact -CustomerId '10000000001' -ContactId '12345'
if ($zepContact.email -ne 'max.mustermann@contoso.com' -or (& (Get-Module InvokeZEP) { $script:ContactCalls.Count }) -ne 1) {
    throw 'Get-ZEPCustomerContact liest den Kontakt nicht oder cacht ihn nicht.'
}
if ($null -ne (Get-ZEPCustomerContact -CustomerId '10000000001' -ContactId '99999')) {
    throw 'Get-ZEPCustomerContact sollte fuer unbekannte Kontakte $null liefern.'
}

Remove-Module InvokeZEP -Force
Import-Module $modulePath -Force

# --- Get-ZEPOfferCategories: single offer, only type 'angebot' by default, empty list without categories
& (Get-Module InvokeZEP) {
    function script:Invoke-ZEPRest {
        param ([string]$Path, [string]$Method = 'GET', [hashtable]$Query, $Body, [int]$RetryCount, [double]$ThrottleDelaySeconds)
        if ($Path -eq 'offers/12345') {
            return [pscustomobject]@{ data = [pscustomobject]@{ id = 12345; categories = @(
                [pscustomobject]@{ name = 'Ausschreibung'; type = [pscustomobject]@{ id = 'angebot'; name = 'Angebot' } },
                [pscustomobject]@{ name = 'Nachtrag'; type = [pscustomobject]@{ id = 'angebot'; name = 'Angebot' } },
                [pscustomobject]@{ name = 'Ausschreibung'; type = [pscustomobject]@{ id = 'projekt'; name = 'Projekt' } }
            ) } }
        }
        return [pscustomobject]@{ data = [pscustomobject]@{ id = 12346; categories = @() } }
    }
}

$offerCategories = @(Get-ZEPOfferCategories -Id 12345)
if (($offerCategories -join ',') -ne 'Ausschreibung,Nachtrag') {
    throw "Get-ZEPOfferCategories liefert falsche Kategorien: $($offerCategories -join ',')"
}
if (@(Get-ZEPOfferCategories -Id 12345 -Type '').Count -ne 2) {
    throw 'Get-ZEPOfferCategories -Type "" sollte alle Kategorien (eindeutige Namen) liefern.'
}
if (@(Get-ZEPOfferCategories -Id 12346).Count -ne 0) {
    throw 'Get-ZEPOfferCategories sollte ohne Kategorien eine leere Liste liefern.'
}

Remove-Module InvokeZEP -Force
Import-Module $modulePath -Force

# --- Get-ZEPProject: path, session cache, $null for unknown projects
& (Get-Module InvokeZEP) {
    $script:ProjectCalls = [System.Collections.Generic.List[string]]::new()
    function script:Invoke-ZEPRest {
        param ([string]$Path, [string]$Method = 'GET', [hashtable]$Query, $Body, [int]$RetryCount, [double]$ThrottleDelaySeconds)
        $script:ProjectCalls.Add($Path)
        if ($Path -eq 'projects/1234') {
            return [pscustomobject]@{ data = [pscustomobject]@{ id = 1234; name = '26B0001'; description = 'Musterprojekt'; dynamicAttributes = @() } }
        }
        $ex = [System.Exception]::new('Not Found')
        Add-Member -InputObject $ex -MemberType NoteProperty -Name Response -Value ([pscustomobject]@{ StatusCode = 404 }) -Force
        throw $ex
    }
}

$zepProject = Get-ZEPProject -Id 1234
$null = Get-ZEPProject -Id 1234
if ($zepProject.name -ne '26B0001' -or (& (Get-Module InvokeZEP) { $script:ProjectCalls.Count }) -ne 1) {
    throw 'Get-ZEPProject liest das Projekt nicht oder cacht es nicht.'
}
if ($null -ne (Get-ZEPProject -Id 9999)) {
    throw 'Get-ZEPProject sollte fuer unbekannte Projekte $null liefern.'
}

Remove-Module InvokeZEP -Force
Import-Module $modulePath -Force

# --- Invoke-ZEPRest waits for Retry-After on HTTP 429 (PS 7 style header and PS 5.1 style header)
& (Get-Module InvokeZEP) {
    function script:Get-ZEPBearerToken { return 'test-token' }
    function script:Start-Sleep { param ([int]$Seconds, [int]$Milliseconds) if ($Seconds) { $script:SleptSeconds += $Seconds } }

    foreach ($headers in @(
        [pscustomobject]@{ RetryAfter = [pscustomobject]@{ Delta = [timespan]::FromSeconds(7); Date = $null } },
        @{ 'Retry-After' = '7' }
    )) {
        $script:SleptSeconds = 0
        $script:RestCalls = 0
        $script:FakeHeaders = $headers
        function script:Invoke-RestMethod {
            param ([uri]$Uri, [string]$Method, [hashtable]$Headers, $ErrorAction)
            $script:RestCalls++
            if ($script:RestCalls -eq 1) {
                $ex = [System.Exception]::new('Too Many Attempts.')
                Add-Member -InputObject $ex -MemberType NoteProperty -Name Response -Value ([pscustomobject]@{ StatusCode = 429; Headers = $script:FakeHeaders }) -Force
                throw $ex
            }
            return [pscustomobject]@{ data = @() }
        }

        $null = Invoke-ZEPRest -Path 'offers' -ThrottleDelaySeconds 0 -WarningAction SilentlyContinue
        if ($script:RestCalls -ne 2 -or $script:SleptSeconds -ne 8) {
            throw "Invoke-ZEPRest wartet bei 429 nicht auf Retry-After (Aufrufe $($script:RestCalls), gewartet $($script:SleptSeconds) s)."
        }
    }
}

Remove-Module InvokeZEP -Force
Import-Module $modulePath -Force

Write-Host 'InvokeZEP module import and export validation passed.' -ForegroundColor Green
