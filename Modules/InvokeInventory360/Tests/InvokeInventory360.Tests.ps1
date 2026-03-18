$moduleRoot = Split-Path -Parent $PSScriptRoot
$manifestPath = Join-Path $moduleRoot 'InvokeInventory360.psd1'

Describe 'InvokeInventory360 module smoke tests' {
    It 'loads the module manifest' {
        $manifest = Test-ModuleManifest $manifestPath

        $manifest.Name | Should Be 'InvokeInventory360'
        $manifest.Version.ToString() | Should Be '1.9.6'
    }

    It 'imports the module and exposes the expected public commands' {
        Remove-Module InvokeInventory360 -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        $commandNames = @(Get-Command -Module InvokeInventory360 | Select-Object -ExpandProperty Name)

        ($commandNames -contains 'Invoke-Inventory') | Should Be $true
        ($commandNames -contains 'Read-PhoneContracts') | Should Be $true
        ($commandNames -contains 'Update-PhoneCards') | Should Be $true
        ($commandNames -contains 'Get-Inventory360Configuration') | Should Be $true
        ($commandNames -contains 'Set-Inventory360Configuration') | Should Be $true
    }

    It 'does not export Pause as a public command' {
        Remove-Module InvokeInventory360 -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        $commandNames = @(Get-Command -Module InvokeInventory360 | Select-Object -ExpandProperty Name)

        ($commandNames -contains 'Pause') | Should Be $false
    }

    It 'updates the exported configuration in-memory' {
        Remove-Module InvokeInventory360 -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        $updatedConfiguration = Set-Inventory360Configuration -ServiceUserName 'TEST_INV' -ApiBaseUri 'https://inventory.example.test/api' -DefaultOrganization 'Example Org' -TelekomInitialContract 'TM999-1' -VodafoneInitialContract 'V999999-1'

        $updatedConfiguration.ServiceUserName | Should Be 'TEST_INV'
        $updatedConfiguration.ApiBaseUri | Should Be 'https://inventory.example.test/api'
        $updatedConfiguration.DefaultOrganization | Should Be 'Example Org'
        $updatedConfiguration.TelekomInitialContract | Should Be 'TM999-1'
        $updatedConfiguration.VodafoneInitialContract | Should Be 'V999999-1'
        $updatedConfiguration.TokenTarget | Should Be 'TEST_INV_ApiToken'
    }

    It 'resets configuration defaults after a fresh import' {
        Remove-Module InvokeInventory360 -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop
        Set-Inventory360Configuration -ServiceUserName 'TEST_INV' -ApiBaseUri 'https://inventory.example.test/api' -DefaultOrganization 'Example Org' -TelekomInitialContract 'TM999-1' -VodafoneInitialContract 'V999999-1' > $null

        Remove-Module InvokeInventory360 -Force -ErrorAction SilentlyContinue
        Import-Module $manifestPath -Force -ErrorAction Stop

        $defaultConfiguration = Get-Inventory360Configuration

        $defaultConfiguration.ServiceUserName | Should Be 'HHPBERLIN_INVENTORY360'
        $defaultConfiguration.ApiBaseUri | Should Be 'https://hhp.enteksystems.de/api/2.0/'
        $defaultConfiguration.DefaultOrganization | Should Be 'hhpberlin - Ingenieure fuer Brandschutz GmbH'
        $defaultConfiguration.TelekomInitialContract | Should Be 'TM058-1'
        $defaultConfiguration.VodafoneInitialContract | Should Be 'V275095-1'
        $defaultConfiguration.TokenTarget | Should Be 'HHPBERLIN_INVENTORY360_ApiToken'
    }
}