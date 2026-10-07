# PowerShell-Modules

This repository contains the public release copies of the following PowerShell modules:

- InvokePersonio
- InvokeInventory360
- InvokeZEP – ZEP offers, offer items and employees (REST API v1)
- InvokeHubspot – HubSpot CRM deals, companies, owners and properties (date-based API version 2026-09)

The published module contents are exported from the validated release artifacts of the internal source workspace.

## Structure

- LICENSE
- Modules/InvokePersonio
- Modules/InvokeInventory360
- Modules/InvokeZEP
- Modules/InvokeHubspot

## Notes

- README files for each module are included inside the module folders.
- Test files are included inside each module folder under Tests.
- Module manifests contain the public metadata used for PowerShell Gallery publication.

## Release Workflow

- Validate and stage releases with `Invoke-ModuleRelease.ps1`.
- The script now uses `Publish-PSResource` as the default PowerShell Gallery publish client.
- The API key can be passed with `-NuGetApiKey`, read from the `NUGET_API_KEY` environment variable or from the Windows Credential Manager.
- Use `-PublishClient PowerShellGet` only when you explicitly need the legacy publish path.