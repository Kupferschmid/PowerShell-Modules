# InvokeInventory360

PowerShell module for reading and updating Inventory360 assets, contracts, and related records via the REST API.

## Features

- Read hardware, SIM, leasing, and phone contract data.
- Create and update Inventory360 records.
- Optional configuration overrides via `Get-Inventory360Configuration` and `Set-Inventory360Configuration`.
- Local credential caching with `BetterCredentials`.
- Azure Automation support via `Get-AutomationPSCredential`.

## Requirements

- PowerShell with the `ImportExcel` module.
- PowerShell with the `BetterCredentials` module version `4.5`.
- Inventory360 API token.

## Public Commands

- `Invoke-Inventory`
- `New-Hardware`
- `New-PhoneContract`
- `New-SimCard`
- `Get-Hardware`
- `Get-PhoneContract`
- `Get-SimCard`
- `Read-Contracts`
- `Read-Leasing`
- `Read-Hardware`
- `Read-PhoneContracts`
- `Read-SimCards`
- `Set-Hardware`
- `Set-PhoneContract`
- `Set-SimCard`
- `Update-PhoneCards`
- `New-PhoneContractFromCsvRecord`
- `New-SimCardsFromCsvRecord`
- `Add-PhoneCardsFromCsvByPhonenumber`
- `Get-NextContractNumber`
- `Format-PhoneNumber`
- `Get-ContractDatesFromCSVRecord`
- `Compare-ShutdownContracts`
- `Import-PhoneContracts`
- `Get-Inventory360Configuration`
- `Set-Inventory360Configuration`

## Credential Behavior

Local sessions:
The module reads the Inventory360 API token from the Windows Credential Manager through `BetterCredentials`.

Azure Automation:
The module reads the API token from `Get-AutomationPSCredential`. There is no local credential fallback in the runbook path.

## Basic Usage

```powershell
Import-Module .\InvokeInventory360.psd1
Get-Inventory360Configuration
Invoke-Inventory -Endpoint 'status'
Read-PhoneContracts
```

## Configuration Example

```powershell
Set-Inventory360Configuration -ServiceUserName 'MY_INVENTORY_APP' -DefaultOrganization 'Example Org'
```

## Tests

Run the smoke tests with Pester:

```powershell
Invoke-Pester -Path .\Tests\InvokeInventory360.Tests.ps1
```