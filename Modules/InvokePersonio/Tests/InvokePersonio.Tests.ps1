$moduleRoot = Split-Path -Parent $PSScriptRoot
$manifestPath = Join-Path $moduleRoot 'InvokePersonio.psd1'

Describe 'InvokePersonio module smoke tests' {
    It 'loads the module manifest' {
        $manifest = Test-ModuleManifest $manifestPath

        $manifest.Name | Should Be 'InvokePersonio'
        $manifest.Version.ToString() | Should Be '1.7.0'
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
            dynamic_1271341 = [pscustomobject]@{ label = 'hhpberlin Kürzel'; value = 'AWE'; type = 'standard'; universal_id = $null }
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
}