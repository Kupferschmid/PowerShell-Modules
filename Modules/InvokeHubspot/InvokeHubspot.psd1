@{
    RootModule = 'InvokeHubspot.psm1'
    ModuleVersion = '0.4.3'
    GUID = '9cb77c4d-3dbd-45c2-8c06-c3f53a9f2e93'
    Author = 'Klaus Kupferschmid'
    CompanyName = 'tempero.it GmbH & hhpberlin GmbH'
    Copyright = '(c) Klaus Kupferschmid. All rights reserved.'
    Description = 'PowerShell module for managing HubSpot CRM deals via REST API with Bearer token authentication.'
    PowerShellVersion = '5.1'
    RequiredModules = @(
        @{
            ModuleName = 'BetterCredentials'
            ModuleVersion = '4.5'
        }
    )
    FunctionsToExport = @(
        'Get-HubspotConfiguration',
        'Set-HubspotConfiguration',
        'Set-HubspotAccessToken',
        'Invoke-HubspotRest',
        'Get-HubspotDeal',
        'Get-HubspotDealDetails',
        'Get-HubspotDeals',
        'Search-HubspotDeals',
        'New-HubspotDeal',
        'Set-HubspotDeal',
        'Remove-HubspotDeal',
        'Get-HubspotOwner',
        'Get-HubspotDealCompanies',
        'Set-HubspotDealPrimaryCompany',
        'Get-HubspotCompany',
        'Get-HubspotProperty',
        'New-HubspotProperty',
        'Add-HubspotPropertyOption',
        'Remove-HubspotDealCompany'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('HubSpot', 'CRM', 'Deals', 'REST', 'PowerShell')
            LicenseUri = 'https://github.com/Kupferschmid/PowerShell-Modules/blob/main/LICENSE'
            ProjectUri = 'https://github.com/Kupferschmid/PowerShell-Modules'
            ReleaseNotes = '0.4.3: Examples, README and tests use placeholder names, e-mail addresses and IDs. 0.4.2: Add-HubspotPropertyOption -UpdateLabel renames existing options. 0.4.1: New Remove-HubspotDealCompany. 0.4.0: Search-HubspotDeals -Operator IN/NOT_IN with -Values, -SortProperty/-SortDescending; new Get-HubspotProperty, New-HubspotProperty, Add-HubspotPropertyOption. 0.3.1: New Get-HubspotCompany (returns $null when the company does not exist). 0.3.0: Migrates all endpoints to date-based HubSpot API version 2026-09 (configurable via -ApiVersion); fixes Windows PowerShell 5.1 (System.Web, json depth); Get-/Set-HubspotDeal -IdProperty; Search-HubspotDeals -Operator (EQ, HAS_PROPERTY, ...); typed write conversion for date/datetime/bool/number/multi-select; date/datetime values read consistently as UTC; new Get-HubspotOwner, Get-HubspotDealCompanies, Set-HubspotDealPrimaryCompany; New-HubspotDeal -PrimaryCompanyId. 0.1.3: Stage/owner display resolution, auth retry, scope-aware warnings, -AllProperties, extended Get-Help examples.'
        }
    }
}
