@{
    RootModule = 'InvokeZEP.psm1'
    ModuleVersion = '0.3.1'
    GUID = '4abf4e3d-44de-4e0a-b4da-6e4a9f8f33c6'
    Author = 'Klaus Kupferschmid'
    CompanyName = 'tempero.it GmbH & hhpberlin GmbH'
    Copyright = '(c) Klaus Kupferschmid. All rights reserved.'
    Description = 'PowerShell module for reading ZEP offers, offer items and employees via REST API.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Get-ZEPConfiguration',
        'Set-ZEPConfiguration',
        'Invoke-ZEPRest',
        'Get-ZEPOffer',
        'Get-ZEPOffers',
        'Get-ZEPOfferItems',
        'Test-ZEPOffersCacheNeedsRefresh',
        'Search-ZEPOffers',
        'ConvertFrom-ZEPDateTime',
        'ConvertTo-ZEPTypedOffer',
        'Get-ZEPRecentOffers',
        'Get-ZEPOffersById',
        'Get-ZEPEmployee'
    )
    CmdletsToExport = @()
    VariablesToExport = '*'
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('ZEP', 'REST', 'Automation', 'PowerShell')
            LicenseUri = 'https://github.com/Kupferschmid/PowerShell-Modules/blob/main/LICENSE'
            ProjectUri = 'https://github.com/Kupferschmid/PowerShell-Modules'
            ReleaseNotes = @'
0.3.1
- Examples and tests use placeholder names and IDs.
0.3.0
- New Get-ZEPRecentOffers: reads only the newest offers (backwards from the last page) instead of all offers.
- New Get-ZEPOffersById: reads several offers in one call via the id[] filter.
- New ConvertFrom-ZEPDateTime: corrects ZEP timestamps (Europe/Berlin wall clock marked as 'Z') to real UTC.
- New ConvertTo-ZEPTypedOffer: reliable types for offers (int ids, string customer ids, UTC datetimes, dates).
- New Get-ZEPEmployee: employee lookup by username with session cache.
- New setting ServerTimeZone (default Europe/Berlin) in Get-/Set-ZEPConfiguration.
- Invoke-ZEPRest supports array query values (repeated keys such as id[]).
- Windows PowerShell 5.1 fixed: System.Web is loaded; offer cache is read without -Depth (was always treated as empty).
0.1.1
- Fixed cache TTL evaluation by handling DateTime/DateTimeOffset values reliably in UTC.
- Improved paging progress logging to include the final page and exact API total.
- Optimized Search-ZEPOffers runtime by avoiding expensive property discovery on valid paths.
- Added optional -SkipPropertyValidation for maximum search performance on known fields.
- Removed obsolete offer-item cache configuration and helper path remnants.
'@
        }
    }
}
