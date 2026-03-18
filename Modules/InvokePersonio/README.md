# InvokePersonio

PowerShell module for reading and updating Personio employee data via the REST API.

## Features

- Read employees and attribute metadata from Personio.
- Update employee attributes through the API.
- Optional configuration overrides via `Get-PersonioConfiguration` and `Set-PersonioConfiguration`.
- Local credential caching with `BetterCredentials`.
- Azure Automation support via `Get-AutomationPSCredential`.

## Requirements

- PowerShell with the `BetterCredentials` module version `4.5`.
- Personio API client credentials.

## Public Commands

- `Invoke-Personio`
- `Get-Employee`
- `Set-Employee`
- `Sync-MobileOwner`
- `Show-Attributes`
- `Get-PersonioConfiguration`
- `Set-PersonioConfiguration`

## Credential Behavior

Local sessions:
The module reads the Personio client credential and cached access token parts from the Windows Credential Manager through `BetterCredentials`.

Azure Automation:
The module reads the Personio client credential from `Get-AutomationPSCredential` and requests a fresh access token for each job. The runbook token is kept in memory only.

## Basic Usage

```powershell
Import-Module .\InvokePersonio.psd1
Get-PersonioConfiguration
Invoke-Personio -Endpoint '?'
Get-Employee -Identity 'user.name'
```

## Configuration Example

```powershell
Set-PersonioConfiguration -ServiceUserName 'MY_PERSONIO_APP' -MailDomain 'example.com'
```

## Tests

Run the smoke tests with Pester:

```powershell
Invoke-Pester -Path .\Tests\InvokePersonio.Tests.ps1
```