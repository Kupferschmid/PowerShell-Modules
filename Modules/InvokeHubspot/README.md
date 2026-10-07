# InvokeHubspot

PowerShell module for managing HubSpot CRM deals using the HubSpot Deals API.

## Scope (step 1)

This first version focuses on CRUD operations for HubSpot deal objects:

- List deals
- Get a single deal by ID
- Create a deal
- Update a deal
- Delete a deal

API reference:
https://developers.hubspot.com/docs/api-reference/latest/crm/objects/deals/guide

## Install / import

```powershell
Import-Module 'c:\Scripts\Powershell\Modules\InvokeHubspot\InvokeHubspot.psd1' -Force
```

## Token setup

On the first call, the module prompts for your HubSpot Private App token, then stores it in Windows Credential Manager via BetterCredentials:

```powershell
Set-HubspotAccessToken -Token 'pat-na1-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
```

You can also pass the token directly with `-AccessToken` on each command.

If no stored token exists, `Invoke-HubspotRest` and the public HubSpot cmdlets prompt once and then persist the token automatically.

The returned deal records are flattened into typed PSCustomObjects: HubSpot property values are promoted to top-level fields, and common types like `bool`, `number`, `date`, `datetime`, `enumeration`, and `json` are converted from the CRM property schema.

## Search

Search a single deal attribute without loading all deals first:

```powershell
Search-HubspotDeals -PropertyName 'dealname' -SearchTerm 'Musterprojekt'
```

The function uses the HubSpot CRM search endpoint and matches token-based partial strings on the selected property.
Use `-Operator` for other comparisons, e.g. exact matches or "property is set":

```powershell
Search-HubspotDeals -PropertyName 'zep_angebots_id' -Operator 'EQ' -SearchTerm '18380'
Search-HubspotDeals -PropertyName 'zep_angebots_id' -Operator 'HAS_PROPERTY' -All
Search-HubspotDeals -PropertyName 'zep_status' -Operator 'IN' -Values '10', '20' -All
Search-HubspotDeals -PropertyName 'zep_angebots_id' -Operator 'HAS_PROPERTY' -SortProperty 'zep_angebots_id' -SortDescending -Limit 1
```

Read or update a deal directly by a unique property instead of the deal ID (no search index delay):

```powershell
Get-HubspotDeal -Id '18380' -IdProperty 'zep_angebots_id'
Set-HubspotDeal -Id '18380' -IdProperty 'zep_angebots_id' -Properties @{ dealname = 'Neuer Name' }
```

## Typed values when writing

`New-HubspotDeal` and `Set-HubspotDeal` format typed PowerShell values according to the HubSpot property type:

| PowerShell value | HubSpot property type | Sent as |
| --- | --- | --- |
| `[datetime]` / `[datetimeoffset]` | `date` | `yyyy-MM-dd` (calendar date of the value) |
| `[datetime]` / `[datetimeoffset]` | `datetime` | ISO 8601 UTC, e.g. `2026-10-07T08:15:00.000Z` |
| `[bool]` | any | `"true"` / `"false"` |
| numbers | any | invariant culture string (`1234.5`) |
| arrays | multi-select enumeration | `value1;value2` |

Strings are passed through unchanged.

## Properties

```powershell
Get-HubspotProperty -Name 'zep_status'                                   # $null if missing
New-HubspotProperty -Definition @{ name = 'demo'; label = 'Demo'; type = 'string'; fieldType = 'text'; groupName = 'dealinformation' }
Add-HubspotPropertyOption -Name 'zep_status' -Value '300' -Label 'Neuer Status'   # keeps existing options
Add-HubspotPropertyOption -Name 'zep_status' -Value '30' -Label 'fertig' -UpdateLabel   # rename an existing option
```

## Owners and companies

```powershell
Get-HubspotOwner -Email 'max.mustermann@contoso.com'          # cached per session
Get-HubspotCompany -Id '10000000001'                         # $null if the company does not exist
Get-HubspotDealCompanies -DealId '12345678903'                # CompanyId, IsPrimary, Labels
Set-HubspotDealPrimaryCompany -DealId '12345678903' -CompanyId '10000000001' -RemovePreviousPrimary
Remove-HubspotDealCompany -DealId '12345678903' -CompanyId '10000000001'      # removes all associations to that company
New-HubspotDeal -Properties @{ dealname = 'Neu' } -PrimaryCompanyId '10000000001'
```

## Configuration

```powershell
Get-HubspotConfiguration
Set-HubspotConfiguration -BaseUri 'https://api.hubapi.com' -ApiVersion '2026-09'
```

## Examples

List deals (first page):

```powershell
Get-HubspotDeals -Limit 10
```

List all deals with paging:

```powershell
Get-HubspotDeals -All -Properties @('dealname','amount','pipeline','dealstage')
```

Get one deal:

```powershell
Get-HubspotDeal -Id '123456789' -Properties @('dealname','amount')
```

Create deal:

```powershell
New-HubspotDeal -Properties @{
    dealname = 'Neuer Deal aus PowerShell'
    amount = '2500'
    closedate = '2026-12-31T00:00:00Z'
    pipeline = 'default'
    dealstage = 'appointmentscheduled'
}
```

Update deal:

```powershell
Set-HubspotDeal -Id '123456789' -Properties @{
    amount = '3000'
    dealstage = 'qualifiedtobuy'
}
```

Delete deal:

```powershell
Remove-HubspotDeal -Id '123456789' -Confirm:$false
```

## Notes

- Ab Version 0.3.0 nutzt das Modul die datumsbasierte HubSpot-API-Version `2026-09` (z. B. `crm/objects/2026-09/deals`). Die alten `v1`-`v4`-APIs werden ab September 2027 nicht mehr unterstuetzt.
- Neue API-Version setzen: `Set-HubspotConfiguration -ApiVersion 'YYYY-MM'` – alle Pfade (Deals, Properties, Pipelines, Owners) werden daraus abgeleitet. `-DealsPath` ueberschreibt nur den Deals-Endpunkt; `-DealsPath ''` stellt die Ableitung wieder her.
- Your sample `curl` can be adapted to this module by using `Get-HubspotDeals -Limit 10`.
- Alle exportierten Cmdlets enthalten erweiterte `Get-Help -Examples` inklusive Troubleshooting-Faellen (z. B. 404/archived, Token-Rotation, Scope-Hinweise).
