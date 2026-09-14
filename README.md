# HelloID-Conn-SA-Sync-EntraID-Owners-To-Resource-Owners

> [!IMPORTANT]
> **Best practice - maximum synchronization frequency: once per week**
>
> Resource ownership changes infrequently, usually weekly or monthly. More frequent
> synchronization causes unnecessary processing without business value, because owner
> changes rarely happen multiple times per week.
>
> If a higher frequency is required for your organization, contact Tools4ever
> Support. This helps determine the appropriate configuration for your use case.

> [!IMPORTANT]
> This repository contains the connector and configuration code only. The implementer
> is responsible for acquiring the connection details, such as tenant ID, application
> ID, certificate and certificate password. Please contact the client's application
> manager to coordinate the connector requirements.

## Table of Contents

- [HelloID-Conn-SA-Sync-EntraID-Owners-To-Resource-Owners](#helloid-conn-sa-sync-entraid-owners-to-resource-owners)
  - [Table of Contents](#table-of-contents)
  - [Requirements](#requirements)
  - [Introduction](#introduction)
  - [Getting started](#getting-started)
    - [HelloID API configuration](#helloid-api-configuration)
    - [App registration and certificate setup](#app-registration-and-certificate-setup)
    - [HelloID-specific configuration](#helloid-specific-configuration)
    - [Convert PFX to base64](#convert-pfx-to-base64)
    - [Synchronization settings](#synchronization-settings)
  - [Remarks](#remarks)
  - [Getting help](#getting-help)
  - [HelloID Docs](#helloid-docs)

## Requirements

- A Microsoft Entra ID app registration with certificate-based authentication.
- Microsoft Graph application permissions to read groups and their owners. At minimum, configure the permissions required by the Graph endpoints used by the script, grant admin consent, and verify them in your tenant.
- Synchronization of Entra ID or local AD users and groups to HelloID. This can be either the [local AD sync](https://docs.helloid.com/en/access-management/directory-sync/active-directory-sync.html) or the [Entra ID sync](https://docs.helloid.com/en/access-management/directory-sync/azure-ad-sync.html).
- The synchronization must be configured to meet your requirements before scheduling.

## Introduction

This connector synchronizes the owners of supported Entra ID groups to the corresponding Resource Owner groups in HelloID. Microsoft 365 groups, including Teams-enabled groups, and eligible security groups can be included.

For every group in scope, the connector queries the owners through Microsoft Graph and adds the matching HelloID users to the Resource Owner group. When `$removeMembers` is enabled, users who are no longer an owner are removed as well.

This connector is intended for situations where Entra ID group ownership is managed outside HelloID. It is designed to work together with the [Entra ID Groups to Products Sync](https://github.com/Tools4everBV/HelloID-Conn-SA-Sync-AzureActiveDirectory-Groups-To-SelfService-Products), which creates the corresponding products and Resource Owner groups.

## Getting started

### HelloID API configuration

1. Go to `Manage portal > Security > API`.
2. Select **Add API key** and create an API key for this scheduled task.
3. Optionally add a descriptive note and restrict the allowed IP addresses.
4. Save the API key.
5. Go to `Manage portal > Automation > Variable library` and verify that the variables used in the [synchronization settings](#synchronization-settings) are available.

### App registration and certificate setup

Create a Microsoft Entra ID app registration for certificate-based, application-only authentication to Microsoft Graph. Follow the official Microsoft documentation for [app-only authentication with a certificate](https://learn.microsoft.com/en-us/entra/identity-platform/certificate-credentials).

Configure the following items in the app registration:

- Create a certificate and upload its public key (`.cer`) under **Certificates & secrets**.
- Store the private key certificate (`.pfx`) securely. The certificate password is required by the scheduled task.
- Add the Microsoft Graph application permissions required to read groups and group owners.
- Grant tenant-wide admin consent for the configured permissions.

The application ID and tenant ID can be found on the **Overview** page of the app registration. The certificate is supplied to HelloID as a base64 string; see the next section.

### HelloID-specific configuration

The scheduled task uses certificate-based authentication and calls the Microsoft Graph `/groups` and `/groups/{id}/owners` endpoints.

Ensure that:

- The app registration has the least-privileged Microsoft Graph application permissions accepted by your tenant for these endpoints.
- Admin consent has been granted.
- The certificate uploaded to Entra ID matches the `.pfx` certificate supplied to HelloID.
- The certificate has not expired and its password is stored as a protected HelloID variable.

For more information, see the Microsoft documentation for [permissions for Microsoft Graph](https://learn.microsoft.com/en-us/graph/permissions-reference) and [list group owners](https://learn.microsoft.com/en-us/graph/api/group-list-owners?view=graph-rest-1.0).

### Convert PFX to base64

HelloID requires the certificate as a base64 string. Use the following example to convert a `.pfx` file:

```powershell
$filePath = 'C:\Cert'
$pfxCertName = 'Cert.pfx'
$pfxPath = "$filePath\$pfxCertName"

$fileContentBytes = [System.IO.File]::ReadAllBytes($pfxPath)
[System.Convert]::ToBase64String($fileContentBytes) | Set-Content "$filePath\HelloID_Cert_Base64.txt"
```

Store the resulting value in the `$EntraIdCertificateBase64String` HelloID variable and the certificate password in `$EntraIdCertificatePassword`.

### Synchronization settings

| Variable name | Description | Notes |
| --- | --- | --- |
| `$portalBaseUrl` | HelloID base URL | Default global variable |
| `$portalApiKey` | HelloID API key | Default global variable |
| `$portalApiSecret` | HelloID API secret | Default global variable |
| `$EntraIdTenantId` | Microsoft Entra ID tenant ID | Recommended as a global variable |
| `$EntraIdAppId` | Microsoft Entra ID application ID | Recommended as a global variable |
| `$EntraIdCertificateBase64String` | Base64 string of the Entra ID application certificate | Recommended as a global variable |
| `$EntraIdCertificatePassword` | Password of the application certificate | Recommended as a protected global variable |
| `$entraIDGroupsSearchFilter` | Optional Microsoft Graph filter for groups to include | When `$null`, all supported groups are queried. Only supported Graph filter properties should be used. See the [Microsoft Graph search query documentation](https://learn.microsoft.com/en-us/graph/search-query-parameter). |
| `$resourceOwnerGroupSource` | Source of the Resource Owner groups in HelloID | If the source is not `Local`, synchronization by the target system can overwrite memberships created by this task. |
| `$resourceOwnerGroupPrefix` | Prefix used to identify the Resource Owner group | Should match the prefix configured by the Entra ID Groups to Products Sync. |
| `$resourceOwnerGroupSuffix` | Suffix used to identify the Resource Owner group | Should match the suffix configured by the Entra ID Groups to Products Sync. |
| `$removeMembers` | Whether to remove users who are no longer owners | When `$true`, manually added users can be removed. Default in the script is `$true`. |

## Remarks

- Review the configured group filter, Resource Owner group prefix and suffix before enabling the scheduled task.
- Enabling `$removeMembers` allows the task to remove existing memberships. Make sure this behavior is desired, because manually added users can be removed.
- This connector is designed to work with the [Entra ID Groups to Products Sync](https://github.com/Tools4everBV/HelloID-Conn-SA-Sync-AzureActiveDirectory-Groups-To-SelfService-Products). Without that synchronization, the corresponding Resource Owner groups may not exist and the task may require changes.
- Test changes with `$dryRun = $true` before performing changes in production.

## Getting help

For more information about configuring a HelloID PowerShell scheduled task, see the [HelloID documentation](https://docs.helloid.com/hc/en-us/articles/115003253294-Create-Custom-Scheduled-Tasks).

## HelloID Docs

The official HelloID documentation is available at [docs.helloid.com](https://docs.helloid.com/).
