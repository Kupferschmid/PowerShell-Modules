$moduleRoot = Split-Path -Parent $PSScriptRoot
$manifestPath = Join-Path $moduleRoot 'InvokeHubspot.psd1'

Describe 'InvokeHubspot module smoke tests' {
    It 'loads the module manifest' {
        $manifest = Test-ModuleManifest $manifestPath

        $manifest.Name | Should Be 'InvokeHubspot'
        $manifest.Version.ToString() | Should Be '0.4.3'
    }

    It 'imports the module and exposes the expected public commands' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        $commandNames = @(Get-Command -Module InvokeHubspot | Select-Object -ExpandProperty Name)

        ($commandNames -contains 'Get-HubspotConfiguration') | Should Be $true
        ($commandNames -contains 'Set-HubspotConfiguration') | Should Be $true
        ($commandNames -contains 'Set-HubspotAccessToken') | Should Be $true
        ($commandNames -contains 'Invoke-HubspotRest') | Should Be $true
        ($commandNames -contains 'Get-HubspotDeal') | Should Be $true
        ($commandNames -contains 'Get-HubspotDealDetails') | Should Be $true
        ($commandNames -contains 'Get-HubspotDeals') | Should Be $true
        ($commandNames -contains 'Search-HubspotDeals') | Should Be $true
        ($commandNames -contains 'New-HubspotDeal') | Should Be $true
        ($commandNames -contains 'Set-HubspotDeal') | Should Be $true
        ($commandNames -contains 'Remove-HubspotDeal') | Should Be $true
        ($commandNames -contains 'Get-HubspotOwner') | Should Be $true
        ($commandNames -contains 'Get-HubspotDealCompanies') | Should Be $true
        ($commandNames -contains 'Set-HubspotDealPrimaryCompany') | Should Be $true
        ($commandNames -contains 'Get-HubspotCompany') | Should Be $true
        ($commandNames -contains 'Get-HubspotProperty') | Should Be $true
        ($commandNames -contains 'New-HubspotProperty') | Should Be $true
        ($commandNames -contains 'Add-HubspotPropertyOption') | Should Be $true
        ($commandNames -contains 'Remove-HubspotDealCompany') | Should Be $true
    }

    It 'updates configuration in-memory' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        (Get-HubspotConfiguration).ApiVersion | Should Be '2026-09'
        (Get-HubspotConfiguration).DealsPath | Should Be 'crm/objects/2026-09/deals'

        $updatedConfiguration = Set-HubspotConfiguration -ServiceUserName 'TEST_HUBSPOT' -BaseUri 'https://api.hubapi.com/' -DealsPath '/crm/objects/2026-09/deals/' -DefaultPageSize 200 -DefaultRetryCount 5 -DefaultThrottleDelaySeconds 0.5

        $updatedConfiguration.ServiceUserName | Should Be 'TEST_HUBSPOT'
        $updatedConfiguration.BaseUri | Should Be 'https://api.hubapi.com'
        $updatedConfiguration.DealsPath | Should Be 'crm/objects/2026-09/deals'
        $updatedConfiguration.DefaultPageSize | Should Be 200
        $updatedConfiguration.DefaultRetryCount | Should Be 5
        $updatedConfiguration.DefaultThrottleDelaySeconds | Should Be 0.5
    }

    It 'derives all CRM paths from the configured API version' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        $configuration = Set-HubspotConfiguration -ApiVersion '2027-03'
        $configuration.ApiVersion | Should Be '2027-03'
        $configuration.DealsPath | Should Be 'crm/objects/2027-03/deals'

        InModuleScope InvokeHubspot {
            Get-HubspotCrmPath -Api 'objects' -Resource 'deals/123' | Should Be 'crm/objects/2027-03/deals/123'
            Get-HubspotCrmPath -Api 'properties' -Resource 'deals' | Should Be 'crm/properties/2027-03/deals'
            Get-HubspotCrmPath -Api 'pipelines' -Resource 'deals' | Should Be 'crm/pipelines/2027-03/deals'
            Get-HubspotCrmPath -Api 'owners' -Resource '87654321' | Should Be 'crm/owners/2027-03/87654321'
        }

        (Set-HubspotConfiguration -DealsPath 'custom/deals').DealsPath | Should Be 'custom/deals'
        (Set-HubspotConfiguration -DealsPath '').DealsPath | Should Be 'crm/objects/2027-03/deals'

        # Pester 3.4 'Should Throw' does not work under PowerShell 7 - check via try/catch instead.
        $threw = $false
        try { $null = Set-HubspotConfiguration -ApiVersion 'v3' -ErrorAction Stop } catch { $threw = $true }
        $threw | Should Be $true
    }

    It 'converts HubSpot records into typed flat objects' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:HubspotPropertyDefinitionsByObjectType = @{
                deals = @{
                    dealname = [pscustomobject]@{ name = 'dealname'; type = 'string' }
                    amount = [pscustomobject]@{ name = 'amount'; type = 'number' }
                    closedate = [pscustomobject]@{ name = 'closedate'; type = 'date' }
                    is_closed = [pscustomobject]@{ name = 'is_closed'; type = 'bool' }
                    stage = [pscustomobject]@{ name = 'stage'; type = 'enumeration'; fieldType = 'checkbox' }
                    custom_json = [pscustomobject]@{ name = 'custom_json'; type = 'json' }
                }
            }

            $record = [pscustomobject]@{
                id = '123'
                archived = $false
                createdAt = '2026-08-06T10:11:12Z'
                updatedAt = '2026-08-06T10:11:13Z'
                properties = [pscustomobject]@{
                    dealname = 'Test Deal'
                    amount = '42.5'
                    closedate = '2026-08-06'
                    is_closed = 'true'
                    stage = 'A;B'
                    custom_json = '{"x":1}'
                }
            }

            $converted = Convert-HubspotRecordToObject -Record $record -ObjectType 'deals'

            $converted.Id | Should Be '123'
            $converted.Dealname | Should Be 'Test Deal'
            $converted.Amount | Should Be 42.5
            $converted.is_closed | Should Be $true
            $converted.stage -join ',' | Should Be 'A,B'
            $converted.custom_json.x | Should Be 1
            ([datetime]$converted.closedate).ToString('yyyy-MM-dd') | Should Be '2026-08-06'
        }
    }

    It 'translates dealstage ids to labels when pipeline metadata is available' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:HubspotPropertyDefinitionsByObjectType = @{
                deals = @{
                    pipeline = [pscustomobject]@{ name = 'pipeline'; type = 'string' }
                    dealstage = [pscustomobject]@{ name = 'dealstage'; type = 'string' }
                }
            }

            $script:HubspotDealPipelineDefinitionsById = @{
                default = [pscustomobject]@{
                    Id = 'default'
                    Label = 'Default Pipeline'
                    Stages = @{
                        contractsent = [pscustomobject]@{
                            Id = 'contractsent'
                            Label = 'Anfrage zugeordnet (hhpberlin Sales)'
                            PipelineId = 'default'
                            PipelineLabel = 'Default Pipeline'
                        }
                    }
                }
            }
            $script:HubspotDealPipelineDefinitionsLoaded = $true

            $record = [pscustomobject]@{
                id = '456'
                archived = $false
                properties = [pscustomobject]@{
                    pipeline = 'default'
                    dealstage = 'contractsent'
                }
            }

            $converted = Convert-HubspotRecordToObject -Record $record -ObjectType 'deals'

            $converted.dealstage | Should Be 'contractsent'
            $converted.dealstageLabel | Should Be 'Anfrage zugeordnet (hhpberlin Sales)'
            ($converted.PSObject.Properties.Name -contains 'dealstageId') | Should Be $false
            $converted.pipeline | Should Be 'default'
            $converted.pipelineLabel | Should Be 'Default Pipeline'
        }
    }

    It 'converts translated dealstage values back to internal ids before writing' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:HubspotDealPipelineDefinitionsById = @{
                default = [pscustomobject]@{
                    Id = 'default'
                    Label = 'Default Pipeline'
                    Stages = @{
                        contractsent = [pscustomobject]@{
                            Id = 'contractsent'
                            Label = 'Anfrage zugeordnet (hhpberlin Sales)'
                            PipelineId = 'default'
                            PipelineLabel = 'Default Pipeline'
                        }
                    }
                }
            }
            $script:HubspotDealPipelineDefinitionsLoaded = $true

            $properties = @{
                pipeline = 'default'
                dealstageId = 'contractsent'
                dealstageLabel = 'Anfrage zugeordnet (hhpberlin Sales)'
            }

            $writeProperties = Convert-HubspotDealPropertiesForWrite -Properties $properties

            $writeProperties.dealstage | Should Be 'contractsent'
            $writeProperties.pipeline | Should Be 'default'
            $writeProperties.Contains('dealstageLabel') | Should Be $false
        }
    }

    It 'returns GUI-like deal details with display labels' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:HubspotPropertyDefinitionsByObjectType = @{
                deals = @{
                    pipeline = [pscustomobject]@{ name = 'pipeline'; label = 'Pipeline'; type = 'enumeration'; fieldType = 'select'; groupName = 'dealinformation'; options = @([pscustomobject]@{ value = 'default'; label = 'hhpberlin Sales' }) }
                    dealstage = [pscustomobject]@{ name = 'dealstage'; label = 'Deal-Phase'; type = 'enumeration'; fieldType = 'select'; groupName = 'dealinformation' }
                    nda_safe = [pscustomobject]@{ name = 'nda_safe'; label = 'NDA / SAFE'; type = 'enumeration'; fieldType = 'radio'; groupName = 'dealinformation'; options = @([pscustomobject]@{ value = 'false'; label = 'Nein' }, [pscustomobject]@{ value = 'true'; label = 'Ja' }) }
                    hubspot_owner_id = [pscustomobject]@{ name = 'hubspot_owner_id'; label = 'Fuer Deal zustaendiger Mitarbeiter'; type = 'number'; fieldType = 'number'; groupName = 'dealinformation' }
                }
            }

            $script:HubspotDealPipelineDefinitionsById = @{
                default = [pscustomobject]@{
                    Id = 'default'
                    Label = 'hhpberlin Sales'
                    Stages = @{
                        '1110595160' = [pscustomobject]@{
                            Id = '1110595160'
                            Label = 'Anfrage zugeordnet'
                            PipelineId = 'default'
                            PipelineLabel = 'hhpberlin Sales'
                        }
                    }
                }
            }
            $script:HubspotDealPipelineDefinitionsLoaded = $true

            function Invoke-HubspotRest {
                param (
                    [Parameter(Mandatory = $false)][string]$Path,
                    [Parameter(Mandatory = $false)][hashtable]$Query,
                    [Parameter(Mandatory = $false)][string]$AccessToken
                )

                return [pscustomobject]@{
                    id = '12345678902'
                    archived = $true
                    createdAt = '2026-08-06T03:28:31Z'
                    updatedAt = '2026-08-06T03:28:35Z'
                    properties = [pscustomobject]@{
                        pipeline = 'default'
                        dealstage = '1110595160'
                        nda_safe = 'false'
                        hubspot_owner_id = '87654321'
                    }
                }
            }

            function Invoke-HubspotRequest {
                param (
                    [Parameter(Mandatory = $false)][string]$ResourcePath,
                    [Parameter(Mandatory = $false)][string]$AccessToken
                )

                if ($ResourcePath -eq 'crm/owners/2026-09/87654321') {
                    return [pscustomobject]@{
                        id = '87654321'
                        firstName = 'Erika'
                        lastName = 'Musterfrau'
                        email = 'erika.musterfrau@example.org'
                    }
                }

                throw "Unexpected resource path: $ResourcePath"
            }

            $details = Get-HubspotDealDetails -Id '12345678902' -Properties @('pipeline', 'dealstage', 'nda_safe', 'hubspot_owner_id')

            ($details | Measure-Object).Count | Should Be 4

            $stageDetail = $details | Where-Object { $_.PropertyName -eq 'dealstage' }
            $stageDetail.Label | Should Be 'Deal-Phase'
            $stageDetail.Value | Should Be '1110595160'
            $stageDetail.DisplayValue | Should Be 'Anfrage zugeordnet'

            $ndaDetail = $details | Where-Object { $_.PropertyName -eq 'nda_safe' }
            $ndaDetail.Label | Should Be 'NDA / SAFE'
            $ndaDetail.Value | Should Be 'false'
            $ndaDetail.DisplayValue | Should Be 'Nein'

            $ownerDetail = $details | Where-Object { $_.PropertyName -eq 'hubspot_owner_id' }
            $ownerDetail.Value | Should Be '87654321'
            $ownerDetail.DisplayValue | Should Be 'Erika Musterfrau'
        }
    }

    It 'retries details read with archived=true on 404' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:HubspotPropertyDefinitionsByObjectType = @{
                deals = @{
                    pipeline = [pscustomobject]@{ name = 'pipeline'; label = 'Pipeline'; type = 'string'; fieldType = 'text'; groupName = 'dealinformation' }
                }
            }

            $script:HubspotDetailsRetryAttempt = 0

            function Invoke-HubspotRest {
                param (
                    [Parameter(Mandatory = $false)][string]$Path,
                    [Parameter(Mandatory = $false)][hashtable]$Query,
                    [Parameter(Mandatory = $false)][string]$AccessToken
                )

                $script:HubspotDetailsRetryAttempt++

                if (-not $Query.ContainsKey('archived')) {
                    $ex = [System.Exception]::new('Not Found')
                    Add-Member -InputObject $ex -MemberType NoteProperty -Name Response -Value ([pscustomobject]@{ StatusCode = 404 }) -Force
                    throw $ex
                }

                return [pscustomobject]@{
                    id = '12345678902'
                    archived = $true
                    properties = [pscustomobject]@{
                        pipeline = 'default'
                    }
                }
            }

            $details = Get-HubspotDealDetails -Id '12345678902' -Properties @('pipeline')

            ($details | Measure-Object).Count | Should Be 1
            $details[0].PropertyName | Should Be 'pipeline'
            $script:HubspotDetailsRetryAttempt | Should Be 2
        }
    }

    It 'keeps explicitly requested empty fields in details output' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:HubspotPropertyDefinitionsByObjectType = @{
                deals = @{
                    nda_safe = [pscustomobject]@{ name = 'nda_safe'; label = 'NDA / SAFE'; type = 'enumeration'; fieldType = 'radio'; groupName = 'dealinformation' }
                }
            }

            function Invoke-HubspotRest {
                param (
                    [Parameter(Mandatory = $false)][string]$Path,
                    [Parameter(Mandatory = $false)][hashtable]$Query,
                    [Parameter(Mandatory = $false)][string]$AccessToken
                )

                return [pscustomobject]@{
                    id = '1'
                    archived = $false
                    properties = [pscustomobject]@{
                        nda_safe = $null
                    }
                }
            }

            $details = Get-HubspotDealDetails -Id '1' -Properties @('nda_safe')
            ($details | Measure-Object).Count | Should Be 1
            $details[0].PropertyName | Should Be 'nda_safe'
            $details[0].Value | Should Be $null
        }
    }

    It 'uses HubSpot standard attributes by default and requests all only with -AllProperties' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:HubspotPropertyDefinitionsByObjectType = @{
                deals = @{
                    dealname = [pscustomobject]@{ name = 'dealname'; type = 'string' }
                    dealstage = [pscustomobject]@{ name = 'dealstage'; type = 'string' }
                    pipeline = [pscustomobject]@{ name = 'pipeline'; type = 'string' }
                }
            }

            $script:CapturedGetDealQuery = $null

            function Invoke-HubspotRest {
                param (
                    [Parameter(Mandatory = $false)][string]$Path,
                    [Parameter(Mandatory = $false)][hashtable]$Query,
                    [Parameter(Mandatory = $false)][string]$AccessToken
                )

                $script:CapturedGetDealQuery = $Query

                return [pscustomobject]@{
                    id = '123'
                    archived = $false
                    properties = [pscustomobject]@{
                        dealname = 'X'
                        dealstage = 'closedlost'
                        pipeline = 'default'
                    }
                }
            }

            $null = Get-HubspotDeal -Id '123'

            $script:CapturedGetDealQuery.ContainsKey('properties') | Should Be $false

            $null = Get-HubspotDeal -Id '123' -AllProperties

            $script:CapturedGetDealQuery.ContainsKey('properties') | Should Be $true
            (($script:CapturedGetDealQuery.properties | Sort-Object) -join ',') | Should Be 'dealname,dealstage,pipeline'
        }
    }

    It 're-prompts token and retries once on invalid authentication' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:TokenPromptCalls = 0
            $script:InvokeRestCalls = 0

            function Get-HubspotBearerToken {
                param (
                    [Parameter(Mandatory = $false)][string]$AccessToken,
                    [Parameter(Mandatory = $false)][switch]$ForcePrompt
                )

                if ($ForcePrompt.IsPresent) {
                    $script:TokenPromptCalls++
                    return 'new-token'
                }

                return 'old-token'
            }

            function Invoke-RestMethod {
                param (
                    [Parameter(Mandatory = $false)][uri]$Uri,
                    [Parameter(Mandatory = $false)][string]$Method,
                    [Parameter(Mandatory = $false)][hashtable]$Headers
                )

                $script:InvokeRestCalls++

                if ($Headers.Authorization -eq 'Bearer old-token') {
                    $ex = [System.Exception]::new('Authentication credentials not found')
                    Add-Member -InputObject $ex -MemberType NoteProperty -Name Response -Value ([pscustomobject]@{ StatusCode = 401 }) -Force
                    $errorRecord = [System.Management.Automation.ErrorRecord]::new($ex, 'INVALID_AUTHENTICATION', [System.Management.Automation.ErrorCategory]::AuthenticationError, $null)
                    $errorRecord.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"status":"error","category":"INVALID_AUTHENTICATION"}')
                    throw $errorRecord
                }

                return [pscustomobject]@{ results = @() }
            }

            $null = Invoke-HubspotRequest -ResourcePath 'crm/objects/2026-09/deals'

            $script:TokenPromptCalls | Should Be 1
            $script:InvokeRestCalls | Should Be 2
        }
    }

    It 're-prompts token when invalid authentication has no status code' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:TokenPromptCalls = 0
            $script:InvokeRestCalls = 0

            function Get-HubspotBearerToken {
                param (
                    [Parameter(Mandatory = $false)][string]$AccessToken,
                    [Parameter(Mandatory = $false)][switch]$ForcePrompt
                )

                if ($ForcePrompt.IsPresent) {
                    $script:TokenPromptCalls++
                    return 'new-token'
                }

                return 'old-token'
            }

            function Invoke-RestMethod {
                param (
                    [Parameter(Mandatory = $false)][uri]$Uri,
                    [Parameter(Mandatory = $false)][string]$Method,
                    [Parameter(Mandatory = $false)][hashtable]$Headers
                )

                $script:InvokeRestCalls++

                if ($Headers.Authorization -eq 'Bearer old-token') {
                    $ex = [System.Exception]::new('Authentication credentials not found. This API supports OAuth 2.0 authentication')
                    $errorRecord = [System.Management.Automation.ErrorRecord]::new($ex, 'INVALID_AUTHENTICATION', [System.Management.Automation.ErrorCategory]::AuthenticationError, $null)
                    $errorRecord.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"status":"error","category":"INVALID_AUTHENTICATION"}')
                    throw $errorRecord
                }

                return [pscustomobject]@{ results = @() }
            }

            $null = Invoke-HubspotRequest -ResourcePath 'crm/objects/2026-09/deals'

            $script:TokenPromptCalls | Should Be 1
            $script:InvokeRestCalls | Should Be 2
        }
    }

    It 'formats typed values according to the HubSpot property type when writing' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:HubspotPropertyDefinitionsByObjectType = @{
                deals = @{
                    abgabe = [pscustomobject]@{ name = 'abgabe'; type = 'date' }
                    seit = [pscustomobject]@{ name = 'seit'; type = 'datetime' }
                }
            }

            $converted = Convert-HubspotDealPropertiesForWrite -Properties @{
                abgabe = [datetime]::new(2026, 10, 14, 0, 0, 0, [DateTimeKind]::Utc)
                seit = [datetime]::new(2026, 10, 6, 23, 24, 35, [DateTimeKind]::Utc)
                nda = $false
                betrag = 1234.5
                id = 18380
                form = @('E-Mail', 'Portal')
                text = 'unveraendert'
            }

            $converted.abgabe | Should Be '2026-10-14'
            $converted.seit | Should Be '2026-10-06T23:24:35.000Z'
            $converted.nda | Should Be 'false'
            $converted.betrag | Should Be '1234.5'
            $converted.id | Should Be '18380'
            $converted.form | Should Be 'E-Mail;Portal'
            $converted.text | Should Be 'unveraendert'
        }
    }

    It 'searches with exact operators and passes idProperty for unique lookups' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:Calls = @()
            function Invoke-HubspotRequest {
                param ([string]$ResourcePath, [string]$Method = 'GET', [hashtable]$Query, $Body, [string]$AccessToken, [int]$RetryCount, [double]$ThrottleDelaySeconds)
                $script:Calls += [pscustomobject]@{ ResourcePath = $ResourcePath; Method = $Method; Query = $Query; Body = $Body }
                return [pscustomobject]@{ id = '1'; properties = [pscustomobject]@{}; results = @() }
            }
            function Convert-HubspotRecordToObject { param ($Record, [string]$AccessToken) return $Record }

            $null = Search-HubspotDeals -PropertyName 'zep_angebots_id' -Operator 'EQ' -SearchTerm '18380'
            $filter = $script:Calls[0].Body.filterGroups[0].filters[0]
            $script:Calls[0].ResourcePath | Should Be 'crm/objects/2026-09/deals/search'
            $filter.operator | Should Be 'EQ'
            $filter.value | Should Be '18380'

            $null = Search-HubspotDeals -PropertyName 'zep_angebots_id' -Operator 'HAS_PROPERTY'
            $script:Calls[1].Body.filterGroups[0].filters[0].ContainsKey('value') | Should Be $false

            $threw = $false
            try { $null = Search-HubspotDeals -PropertyName 'dealname' -Operator 'EQ' -ErrorAction Stop } catch { $threw = $true }
            $threw | Should Be $true

            $null = Get-HubspotDeal -Id '18380' -IdProperty 'zep_angebots_id'
            $script:Calls[2].ResourcePath | Should Be 'crm/objects/2026-09/deals/18380'
            $script:Calls[2].Query.idProperty | Should Be 'zep_angebots_id'
        }
    }

    It 'resolves owners by email with session cache' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:OwnerCalls = 0
            function Invoke-HubspotRequest {
                param ([string]$ResourcePath, [string]$Method = 'GET', [hashtable]$Query, $Body, [string]$AccessToken)
                $script:OwnerCalls++
                $ResourcePath | Should Be 'crm/owners/2026-09'
                $Query.email | Should Be 'max.mustermann@contoso.com'
                return [pscustomobject]@{ results = @([pscustomobject]@{ id = '12345678'; email = 'max.mustermann@contoso.com'; firstName = 'Max'; lastName = 'Mustermann'; userId = 12345678; archived = $false }) }
            }

            (Get-HubspotOwner -Email 'max.mustermann@contoso.com').Id | Should Be '12345678'
            (Get-HubspotOwner -Email ' Max.Mustermann@contoso.com ').Id | Should Be '12345678'
            $script:OwnerCalls | Should Be 1
        }
    }

    It 'sets the primary company and removes the previous primary company' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:Calls = @()
            function Invoke-HubspotRequest {
                param ([string]$ResourcePath, [string]$Method = 'GET', [hashtable]$Query, $Body, [string]$AccessToken, [int]$RetryCount, [double]$ThrottleDelaySeconds)
                $script:Calls += [pscustomobject]@{ ResourcePath = $ResourcePath; Method = $Method; Body = $Body }
                if ($Method -eq 'GET') {
                    return [pscustomobject]@{ results = @(
                        [pscustomobject]@{ toObjectId = 111; associationTypes = @([pscustomobject]@{ typeId = 5; label = 'Primary' }, [pscustomobject]@{ typeId = 341; label = $null }) },
                        [pscustomobject]@{ toObjectId = 333; associationTypes = @([pscustomobject]@{ typeId = 341; label = $null }) }
                    ) }
                }
            }

            $companies = @(Get-HubspotDealCompanies -DealId '42')
            $companies.Count | Should Be 2
            ($companies | Where-Object IsPrimary).CompanyId | Should Be '111'

            $script:Calls = @()
            Set-HubspotDealPrimaryCompany -DealId '42' -CompanyId '222' -RemovePreviousPrimary
            $put = $script:Calls | Where-Object Method -eq 'PUT'
            $put.ResourcePath | Should Be 'crm/objects/2026-09/deals/42/associations/companies/222'
            (@($put.Body) | ForEach-Object { $_.associationTypeId }) -join ',' | Should Be '5,341'
            ($script:Calls | Where-Object Method -eq 'DELETE').ResourcePath | Should Be 'crm/objects/2026-09/deals/42/associations/companies/111'
        }
    }

    It 'searches with IN values and sorting' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:Calls = @()
            function Invoke-HubspotRequest {
                param ([string]$ResourcePath, [string]$Method = 'GET', [hashtable]$Query, $Body, [string]$AccessToken, [int]$RetryCount, [double]$ThrottleDelaySeconds)
                $script:Calls += [pscustomobject]@{ Body = $Body }
                return [pscustomobject]@{ results = @() }
            }

            $null = Search-HubspotDeals -PropertyName 'zep_status' -Operator 'IN' -Values '10', '20' -SortProperty 'zep_angebots_id' -SortDescending
            $filter = $script:Calls[0].Body.filterGroups[0].filters[0]
            $filter.operator | Should Be 'IN'
            (@($filter.values) -join ',') | Should Be '10,20'
            $filter.ContainsKey('value') | Should Be $false
            $script:Calls[0].Body.sorts[0].propertyName | Should Be 'zep_angebots_id'
            $script:Calls[0].Body.sorts[0].direction | Should Be 'DESCENDING'

            $threw = $false
            try { $null = Search-HubspotDeals -PropertyName 'zep_status' -Operator 'IN' -ErrorAction Stop } catch { $threw = $true }
            $threw | Should Be $true
        }
    }

    It 'adds a property option while keeping existing options' {
        Remove-Module InvokeHubspot -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        InModuleScope InvokeHubspot {
            $script:Patch = $null
            function Invoke-HubspotRequest {
                param ([string]$ResourcePath, [string]$Method = 'GET', [hashtable]$Query, $Body, [string]$AccessToken, [int]$RetryCount, [double]$ThrottleDelaySeconds)
                if ($Method -eq 'PATCH') { $script:Patch = [pscustomobject]@{ ResourcePath = $ResourcePath; Body = $Body }; return $null }
                return [pscustomobject]@{ name = 'zep_status'; options = @([pscustomobject]@{ value = '10'; label = 'neu'; displayOrder = 0; hidden = $false }) }
            }

            $null = Add-HubspotPropertyOption -Name 'zep_status' -Value '10' -Label 'neu'
            $script:Patch | Should Be $null

            $null = Add-HubspotPropertyOption -Name 'zep_status' -Value '300' -Label 'Neu in ZEP'
            $script:Patch.ResourcePath | Should Be 'crm/properties/2026-09/deals/zep_status'
            (@($script:Patch.Body.options | ForEach-Object { $_.value }) -join ',') | Should Be '10,300'
            $script:Patch.Body.options[1].displayOrder | Should Be 1

            $script:Patch = $null
            $null = Add-HubspotPropertyOption -Name 'zep_status' -Value '10' -Label 'neu (ZEP)'
            $script:Patch | Should Be $null

            $null = Add-HubspotPropertyOption -Name 'zep_status' -Value '10' -Label 'neu (ZEP)' -UpdateLabel
            @($script:Patch.Body.options).Count | Should Be 1
            $script:Patch.Body.options[0].value | Should Be '10'
            $script:Patch.Body.options[0].label | Should Be 'neu (ZEP)'
        }
    }
}
