# sharepoint-online-audit
PowerShell scripts to audit SharePoint Online sites

## Files
| File | Role |
| - | - |
| **_Helpers.ps1** | Useful methods |
| **Config.Lab.json** | Configuration file with tenant and SharePoint URL, app registration and CSV separator |
| **WebPartId.csv** | List of SharePoint modern webparts with Name and GUID |
| **GetSharePointWebParts.ps1** | Script to get webpart used in SharePoint sites |
| **SetDisabledWebPartIds.ps1** | Script to disable SharePoint webparts |
| **GetSearchManagedProperties.ps1** | Script to get search managed properties |

## Prerequisities

**Pnp PowerShell** must be installed on your computer.

To check if the module is installed, you can use the following command:

~~~powershell
Get-Module PnP.PowerShell -ListAvailable | Select-Object Name,Version | Sort-Object Version -Descending 
~~~
If not, you can install the latest version:

~~~powershell
Install-Module PnP.PowerShell 
~~~

## Configuration

Create a new configuration file based on `Config.LAB.json` or edit this one.

When executing the scripts, the code-name of the configuration should be passed as an argument:

~~~powershell
.\GetSharePointWebParts.ps1 -Env LAB
.\GetSharePointWebParts.ps1 -Env PROD
~~~

To connect to SharePoint, it's better to have an **Azure app registration** with **Sites.FullControl.All** application permission (key `ClientId`). The app registration must used a certificate. This certificate can be referenced by its path and password (keys `PfxPath` and `PfxPwd`) or installed  in your certifcates store and referenced by its thumbprint (key: `CertThumb`). If no certificate with password or thumbprint is indicated in the configuration file, the connection will use your browser (interactive mode).

~~~json
{
    "OutputDir": "C:/Temp",
    "Sep" : ";",
    "ListUrl": "SitePages",
    "WebPartIdCsv": "WebPartId.csv",
    "ClientId" : "guid-of-your-app-registration or empty",
    "TenantSiteUrl": "https://your-tenant-admin.sharepoint.com/",
    "CertThumb": "your-pfx-thumbprint or empty",
    "PfxPath": "C:/Temp/your-pfx.pfx or empty",
    "PfxPwd": "your-pfx-password or empty",
    "TenantName": "your-tenant.onmicrosoft.com"
}
~~~

## Disable webparts

The separator shoud be the same, as indicated in the configuration file (comma, semi-colon, etc.).

Entrer the name to one or many webparts to disable tenant-wide.

To disable nothing (so, to autorize everything), just entrer an empty value `""`

    .\SetDisabledWebPartIds.ps1 -Env LAB -WebPartNames "AmazonKindle;Twitter"

    .\SetDisabledWebPartIds.ps1 -Env LAB -WebPartNames ""

You can add `-Verbose` to display more information in the terminal.

    .\SetDisabledWebPartIds.ps1 -Env LAB -WebPartNames "AmazonKindle;Twitter" -Verbose

By design, **only Kindle/YouTube/Twitter/ContentEmbed/VideoEmbed/MicrosoftBookings web parts can be disabled**.
If you try to disable other webpart, such as Office 365 Connectors, you will have an error!

## Export webparts

The script has three steps:
1. Get modern SharePoint sites by filter the template and generate a CSV file
2. For each site in the CSV file, open every modern pages and get webparts definition
3. Generate a CSV file with webparts, pages and sites

The CSV files are generated in the folder specified in the configuration file (`OutputDir`):
- SharePointUrls-your-tenant.onmicrosoft.com.csv
- SharePointWebParts-your-tenant.onmicrosoft.com.csv

~~~powershell
.\GetSharePointWebParts.ps1 -Env LAB
~~~

You can add `-Verbose` to display more information in the terminal.

~~~powershell
.\GetSharePointWebParts.ps1 -Env LAB -Verbose
~~~

## Get search managed properties

The script returns the search managed properties configured at the tenant level.

The result is an array:

| Name | Aliases | Mappings | Type |
| - | - | - | - |
| RefinableString01 | {Originator} | {ows_Originator} | Text |
| RefinableString02 | {} | {ows_TeamType} | Text |

~~~powershell
.\GetSearchManagedProperties.ps1 -Env LAB
~~~

You can add `-Verbose` to display more information in the terminal.

~~~powershell
.\GetSearchManagedProperties.ps1 -Env LAB -Verbose
~~~