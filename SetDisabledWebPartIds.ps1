[CmdletBinding()]
param(
	[string]$Env = "LAB",
	[string]$WebPartNames = "AmazonKindle;Twitter"
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

	Write-Host "Import-Csv $($config.WebPartIdCsv)" -NoNewline:$True
	$webpartsId = Import-Csv $($config.WebPartIdCsv) -Delimiter $($config.Sep) -Encoding UTF8
	Write-Host " [OK]"  -ForegroundColor Green

	$webpartsToDisable = @()
	$webparts = $WebPartNames -split $config.Sep
	foreach ($webpart in $webparts) {
		$wid = ($webpartsId | Where-Object -Property WebPartName -eq $webpart).WebPartId
		if ($null -ne $wid) {
			$webpartsToDisable += $wid
			Write-Verbose "$wid is the GUID for the webpart [$webpart]"
		}
		else {
			Write-Verbose "Unable to find GUID for the webpart [$webpart]"
		}
	}

	try {
		Connect-Site $config "Admin" $config.tenantSiteUrl

		Write-Host "Set-PnPTenant -DisabledWebPartIds $webpartsToDisable" -NoNewline:$True

		# only Kindle/YouTube/Twitter/ContentEmbed/VideoEmbed/MicrosoftBookings web parts can be disabled now
		Set-PnPTenant -DisabledWebPartIds $webpartsToDisable

		Write-Host " [OK]"  -ForegroundColor Green
	}
	catch {
		Write-Error $_
	}
	finally {
		Disconnect-Site
	}
}
catch {
	Write-Error $_
}
finally {
	Invoke-Stop
}