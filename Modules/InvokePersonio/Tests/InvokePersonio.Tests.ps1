$moduleRoot = Split-Path -Parent $PSScriptRoot
$manifestPath = Join-Path $moduleRoot 'InvokePersonio.psd1'

Describe 'InvokePersonio module smoke tests' {
    It 'loads the module manifest' {
        $manifest = Test-ModuleManifest $manifestPath

        $manifest.Name | Should Be 'InvokePersonio'
        $manifest.Version.ToString() | Should Be '1.9.0'
    }

    It 'imports the module and exposes the expected public commands' {
        Remove-Module InvokePersonio -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        $commandNames = @(Get-Command -Module InvokePersonio | Select-Object -ExpandProperty Name)

        ($commandNames -contains 'Invoke-Personio') | Should Be $true
        ($commandNames -contains 'Get-Employee') | Should Be $true
        ($commandNames -contains 'Set-Employee') | Should Be $true
        ($commandNames -contains 'Sync-MobileOwner') | Should Be $true
        ($commandNames -contains 'Show-Attributes') | Should Be $true
        ($commandNames -contains 'Get-PersonioConfiguration') | Should Be $true
        ($commandNames -contains 'Set-PersonioConfiguration') | Should Be $true
    }

    It 'updates the exported configuration in-memory' {
        Remove-Module InvokePersonio -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        $updatedConfiguration = Set-PersonioConfiguration -ServiceUserName 'TEST_USER' -BaseUri 'https://example.test/v1/company/employees/' -AuthUri 'https://example.test/v1/auth?' -MailDomain '@example.test'

        $updatedConfiguration.ServiceUserName | Should Be 'TEST_USER'
        $updatedConfiguration.BaseUri | Should Be 'https://example.test/v1/company/employees'
        $updatedConfiguration.AuthUri | Should Be 'https://example.test/v1/auth?'
        $updatedConfiguration.MailDomain | Should Be 'example.test'
        $updatedConfiguration.AccessToken1 | Should Be 'TEST_USER_PersonioAccessToken_1'
        $updatedConfiguration.AccessToken2 | Should Be 'TEST_USER_PersonioAccessToken_2'
    }

    It 'converts raw Personio attribute objects into scalar user properties' {
        Remove-Module InvokePersonio -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        $rawEmployee = [pscustomobject]@{
            id = [pscustomobject]@{ label = 'ID'; value = 1807377; type = 'integer'; universal_id = 'id' }
            first_name = [pscustomobject]@{ label = 'First name'; value = 'Astrid'; type = 'standard'; universal_id = 'first_name' }
            hire_date = [pscustomobject]@{ label = 'Hire date'; value = '2009-07-01T00:00:00+02:00'; type = 'date'; universal_id = 'hire_date' }
            dynamic_1271341 = [pscustomobject]@{ label = 'hhpberlin Kuerzel'; value = 'AWE'; type = 'standard'; universal_id = $null }
            office = [pscustomobject]@{ label = 'Workplace'; value = [pscustomobject]@{ type = 'Office'; attributes = [pscustomobject]@{ name = 'Berlin' } }; type = 'standard'; universal_id = 'office' }
        }

        $user = & (Get-Module InvokePersonio) {
            param($employee)
            ConvertTo-UserObject -employees $employee
        } $rawEmployee

        $user.id | Should Be 1807377
        $user.first_name | Should Be 'Astrid'
        $user.hire_date.GetType().Name | Should Be 'DateTime'
        $user.hhpberlin_Kuerzel | Should Be 'AWE'
        $user.office.type | Should Be 'Office'
    }

    It 'resets configuration defaults after a fresh import' {
        Remove-Module InvokePersonio -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop
        Set-PersonioConfiguration -ServiceUserName 'TEST_USER' -BaseUri 'https://example.test/v1/company/employees/' -AuthUri 'https://example.test/v1/auth?' -MailDomain '@example.test' > $null

        Remove-Module InvokePersonio -Force -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        $defaultConfiguration = Get-PersonioConfiguration

        $defaultConfiguration.ServiceUserName | Should Be 'HHPBERLIN_USERMANAGER'
        $defaultConfiguration.BaseUri | Should Be 'https://api.personio.de/v1/company/employees'
        $defaultConfiguration.AuthUri | Should Be 'https://api.personio.de/v1/auth?'
        $defaultConfiguration.MailDomain | Should Be 'hhpberlin.de'
        $defaultConfiguration.AccessToken1 | Should Be 'HHPBERLIN_USERMANAGER_PersonioAccessToken_1'
        $defaultConfiguration.AccessToken2 | Should Be 'HHPBERLIN_USERMANAGER_PersonioAccessToken_2'
    }

    It 'uses the email endpoint instead of the paged list endpoint' {
        Remove-Module InvokePersonio -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop
        Set-PersonioConfiguration -BaseUri 'https://example.test/v1/company/employees/' -MailDomain '@example.test' > $null

        $result = & (Get-Module InvokePersonio) {
            $script:lastUri = $null

            function Get-Creds {
                @(
                    [pscustomobject]@{
                        Password = ConvertTo-SecureString -String 'token-part-1' -AsPlainText -Force
                    }
                )
            }

            function Remove-StoredCredentialSafe {}
            function Set-Credential {}
            function Write-Host {}
            function Invoke-WebRequest {
                param($Uri, $Method, $Headers, $Body, [switch]$UseBasicParsing, $ErrorAction)

                $script:lastUri = $Uri
                [pscustomobject]@{
                    Content = '{"Data":{"attributes":{"id":{"label":"ID","value":1807377,"type":"integer","universal_id":"id"},"email":{"label":"Email","value":"c.abel@example.test","type":"standard","universal_id":"email"}}}}'
                    Headers = @{ authorization = 'Bearer token-part-1' }
                }
            }

            $response = Invoke-Personio '?email=c.abel%40example.test' -RawOutput

            [pscustomobject]@{
                Uri = $script:lastUri
                Count = $response.Count
            }
        }

        $result.Uri | Should Be 'https://example.test/v1/company/employees?email=c.abel%40example.test'
        $result.Count | Should Be 1
    }

    It 'shows only the yellow token refresh message after a 401 retry' {
        Remove-Module InvokePersonio -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop
        Set-PersonioConfiguration -BaseUri 'https://example.test/v1/company/employees/' -MailDomain '@example.test' > $null

        $result = & (Get-Module InvokePersonio) {
            $script:webRequestCallCount = 0
            $script:renewCalled = $false
            $script:hostMessages = @()

            function Get-Creds {
                param (
                    $connectionTarget = 'PER',
                    [boolean] $renew = $false
                )

                if ($renew) {
                    $script:renewCalled = $true
                }

                @(
                    [pscustomobject]@{
                        Password = ConvertTo-SecureString -String 'token-part-1' -AsPlainText -Force
                    }
                )
            }

            function Remove-StoredCredentialSafe {}
            function Set-Credential {}
            function Write-Host {
                param($Object, $ForegroundColor)

                $script:hostMessages += [pscustomobject]@{
                    Object = $Object
                    ForegroundColor = [string] $ForegroundColor
                }
            }
            function Invoke-WebRequest {
                param($Uri, $Method, $Headers, $Body, [switch]$UseBasicParsing, $ErrorAction)

                $script:webRequestCallCount++

                if ($script:webRequestCallCount -eq 1) {
                    throw 'WebRequest-Error: 401'
                }

                [pscustomobject]@{
                    Content = '{"Data":{"attributes":{"id":{"label":"ID","value":1807377,"type":"integer","universal_id":"id"},"email":{"label":"Email","value":"c.abel@example.test","type":"standard","universal_id":"email"}}}}'
                    Headers = @{ authorization = 'Bearer token-part-1' }
                }
            }

            $response = Get-Employee -identity 'c.abel'

            [pscustomobject]@{
                ResultEmail = $response.email
                RenewCalled = $script:renewCalled
                Messages = @($script:hostMessages)
            }
        }

        @($result.Messages | Where-Object Object -eq 'Access-Token wurde erneuert').Count | Should Be 1
        @($result.Messages | Where-Object Object -eq 'personio_token ist falsch WebRequest-Error: 401').Count | Should Be 0
        $result.RenewCalled | Should Be $true
        $result.ResultEmail | Should Be 'c.abel@example.test'
    }
}