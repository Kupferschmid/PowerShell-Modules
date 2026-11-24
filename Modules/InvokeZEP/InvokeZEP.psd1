@{
    RootModule = 'InvokeZEP.psm1'
    ModuleVersion = '0.7.0'
    GUID = '4abf4e3d-44de-4e0a-b4da-6e4a9f8f33c6'
    Author = 'Klaus Kupferschmid'
    CompanyName = 'tempero.it GmbH & hhpberlin GmbH'
    Copyright = '(c) Klaus Kupferschmid. All rights reserved.'
    Description = 'PowerShell module for reading ZEP offers, offer items, offer categories, offer documents, projects, employees and customer contacts via REST API (offer documents via SOAP).'
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
        'Get-ZEPOffersByStatus',
        'Get-ZEPOffersById',
        'Get-ZEPEmployee',
        'Get-ZEPCustomerContact',
        'Get-ZEPOfferCategories',
        'Get-ZEPProject',
        'Invoke-ZEPSoap',
        'Get-ZEPOfferDocuments'
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
0.7.0
- New Invoke-ZEPSoap: calls the ZEP SOAP interface (V10) with its own authorization token (Credential Manager target SoapTokenTarget, asked for once in interactive sessions like the REST token; in Azure Automation the credential SoapServiceUserName; never prompts in non-interactive sessions).
- New Get-ZEPOfferDocuments: documents of an offer with the currently valid version and creation time (SOAP readDokumente; the REST API has no version information).
- New settings SoapUri, SoapTokenTarget, SoapServiceUserName in Get-/Set-ZEPConfiguration.
- Invoke-ZEPSoap raises SOAP faults and response headers with returnCode <> 0 as errors (ZEP fault text included), waits for Retry-After on HTTP 429/503 and retries other server errors briefly, like Invoke-ZEPRest.
0.6.1
- Help of Get-ZEPProject uses a placeholder project number.
0.6.0
- New Get-ZEPOfferCategories: category names of an offer (GET /offers/{id}, ZEP REST API 1.8.0), by default only type 'angebot'.
- New Get-ZEPProject: reads a project by id (e.g. the project_id of an offer; project number = name, dynamic attributes included) with session cache.
0.5.1
- HTTP 429: waits for the Retry-After time sent by ZEP (default 60 s) instead of retrying after 2/4 s (ZEP limit: 1000 requests per 5 minutes per tenant).
0.5.0
- New Get-ZEPCustomerContact: reads a contact person of a customer (e.g. the customer_contact_id of an offer) with session cache.
0.4.0
- New Get-ZEPOffersByStatus: reads all offers with the given status ids via orderBy=status (supported by ZEP since 23.10.2026) and stops after the highest requested status. Throws if the API returns an unsorted list.
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
