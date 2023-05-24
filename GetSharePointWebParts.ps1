[CmdletBinding()]
param(
	[string]$Env = "LAB"
)

$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
try {
	. ("$ScriptDirectory\_Helpers.ps1")
}
catch {
	Write-Error "Error while loading PowerShell scripts" 
	Write-Error $_.Exception.Message
}

Invoke-Start $MyInvocation.MyCommand.Name $ScriptDirectory

try {
	$config = Get-Config $Env
	$config
	$SharePointUrlCsv = "$($config.OutputDir)\SharePointUrls-$($config.TenantName).csv"

	try {
		Connect-Site $config "Admin" $config.tenantSiteUrl
	
		Write-Host "Get-PnPTenantSite and Export-CSV to $SharePointUrlCsv" -NoNewline:$True
		Get-PnPTenantSite -Detailed `
		| Where-Object -Property Template -In ("GROUP#0", "STS#3", "SITEPAGEPUBLISHING#0", "TEAMCHANNEL#0") `
		| Select-Object Title, Url, Template `
		| Export-CSV $SharePointUrlCsv -NoTypeInformation -Delimiter $($config.Sep) -Encoding UTF8
		Write-Host " [OK]"  -ForegroundColor Green
	}
	catch {
		Write-Warning $_
	}
	finally {
		Disconnect-Site
	}

	Write-Host "Import-Csv $($config.WebPartIdCsv)" -NoNewline:$True
	$webpartsId = Import-Csv $($config.WebPartIdCsv) -Delimiter $($config.Sep) -Encoding UTF8
	Write-Host " [OK]"  -ForegroundColor Green

	$data = @()
	$urls = Import-Csv $SharePointUrlCsv -Delimiter $($config.Sep) -Encoding UTF8
	if ($null -ne $urls -and $urls.Length -gt 0) {
		$i = 0
		$increment = 100 / $urls.Length

		foreach ($site in $urls) {
			$siteUrl = $site.Url

			$i += 1
			$ip = [int]($increment * $i)
			Write-Progress -Activity "Processing data" -Status "$ip%" -PercentComplete $ip

			try {
				Connect-Site $config $site.Title $siteUrl
			
				$pages = $null
				try {
					$listPages = Get-PnPList | Where-Object { $_.RootFolder.ServerRelativeUrl -like "*$($config.ListUrl)" } | Select-Object -First 1
					if ($null -eq $listPages) {
						Write-Warning "List [$($config.ListUrl)] does not exist for site $siteUrl"
					}
					else {
						$pages = Get-PnPListItem -List $listPages
					}
				}
				catch {
					Write-Warning $_
					continue
				}

				if ($null -ne $pages) {
					foreach ($page in $pages) {
						$pageUrl = $page.FieldValues.FileLeafRef
						try {
							$clientPage = Get-PnPPage -Identity $pageUrl -ErrorAction SilentlyContinue
							Write-Verbose "  - $pageUrl"
							if ($null -ne $clientPage) {
								foreach ($webpart in $clientPage.Controls) {	
									Write-Verbose "   -  $($webpart.Title)"
									$webpartJson = $null
									if ($null -ne $($webpart.PropertiesJson)) {
										try {
											$webpartJson = $webpart.PropertiesJson | ConvertFrom-Json
										}
										catch {
											Write-Host -ForegroundColor DarkMagenta $_
										}
									} 
									$obj = [PSCustomObject]@{
										"Site"                   = $site.Title
										"Page"                   = $clientPage.PageTitle
										"URL"                    = "$siteUrl/$($config.ListUrl)/$($clientPage.Name)"
										"WebPart Title"          = $webpart.Title    
										"WebPart InstanceId"     = $webpart.InstanceId 
										"WebPart WebPartId"      = $webpart.WebPartId 
										"WebPart WebPartType"    = ($webpartsId | Where-Object -Property WebPartId -eq $webpart.WebPartId).WebPartName
										"WebPart ProviderName"   = $webpartJson.profiles.providerName -join $($config.Sep) 
										"WebPart PropertiesJson" = $webpart.PropertiesJson 
									}
									$data += $obj
								}
							}
						}
						catch {
							Write-Error $_
						}
					}
				}
			}
			catch {
				Write-Error $_
			}
		}

		$SharePointWebpartCsv = "$($config.OutputDir)\SharePointWebParts-$($config.TenantName).csv"
		Write-Host "Generate file $SharePointWebpartCsv" -NoNewline
		$data | Export-CSV $SharePointWebpartCsv -NoTypeInformation -Delimiter $($config.Sep) -Encoding UTF8
		Write-Host -ForegroundColor Green " [OK]"
	}
}
catch {
	Write-Error $_
}
finally {
	Invoke-Stop
}