####### This scirpt is used to force a full synchronization of your on prem users password with office 365######## 
#####Open Powershell on the AAD sync server with admin privileges and run the following script. Kindly replace the "Domain.com" with your On prem MIIS connector name and “TenantName.onmicrosoft.com – AAD” with your MIIS Azure connector name ######## 
 
 
$Local = "Domain.com" 
 
$Remote = "TenantName.onmicrosoft.com - AAD" 
 
#Import Azure Directory Sync Module to Powershell 
 
Import-Module AdSync 
 
$OnPremConnector = Get-ADSyncConnector -Name $Local 
 
Write-Output "On Prem Connector information received" 
 
$Object = New-Object Microsoft.IdentityManagement.PowerShell.ObjectModel.ConfigurationParameter "Microsoft.Synchronize.ForceFullPasswordSync", String, ConnectorGlobal, $Null, $Null, $Null 
 
$Object.Value = 1 
 
$OnPremConnector.GlobalParameters.Remove($Object.Name) 
 
$OnPremConnector.GlobalParameters.Add($Object) 
 
$OnPremConnector = Add-ADSyncConnector -Connector $OnPremConnector 
 
Set-ADSyncAADPasswordSyncConfiguration -SourceConnector $Local -TargetConnector $Remote -Enable $False 
 
Set-ADSyncAADPasswordSyncConfiguration -SourceConnector $Local -TargetConnector $Remote -Enable $True