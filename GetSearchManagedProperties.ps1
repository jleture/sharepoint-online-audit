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

	Connect-Site $config "Admin" $config.tenantSiteUrl

	Write-Host "Get-PnPSearchConfiguration ManagedPropertyMappings"

	Get-PnPSearchConfiguration -Scope Subscription -OutputFormat ManagedPropertyMappings | Select-Object -Property * | Format-Table
}
catch {
	Write-Error $_
}
finally {
	Invoke-Stop
}