<# 
.SYNOPSIS  
    Connect to SharePoint Online (SPO) Admin Site and display Urls of all provisioned site collections.
    
.DESCRIPTION 
    This runbook displays Url of all the provisioned SPO site collections forparticular Office 365 tenant.
Prereq:    
    1. Automation Module Asset containing SPO integration module.
    2. Automation Credential Asset containing the admin user id and the password for SPO tenant. 

.PARAMETER  SPOAdminSiteUrl 
    String Url to the SPO Admin Site. 
    Example: https//:{tenant}-admin.sharepoint.com 
 
.PARAMETER  PSCredName
    String name of PSCredential Asset. 
    Example: SPOAdminCred. The asset with name 'SPOAdminCred'' must be present as a Credential asset of type PSCredential. 

.EXAMPLE 
   Display-AllSPOSites -SPOAdminSiteUrl "https://{tenant}-admin.sharepoint.com" -PSCredName "SPOAdminCred"

.NOTES 
    Author: Razi Rais 
    Details: http://tinyurl.com/SPOAzAutomate  
    Last Updated: 8/22/2014    
#> 
workflow Display-AllSPOSites
{
    param( 
            [OutputType([string[]])]
          
            [Parameter(Mandatory=$true)]
            [string] 
            $SPOAdminSiteUrl,
            
            [Parameter(Mandatory=$true)]   
            [string] 
            $PSCredName
         )
        
 $PSCred = Get-AutomationPSCredential -Name $PSCredName
 Write-Output "Connecting to SPO Admin Site: '$SPOAdminSiteUrl' using account:" $PSCred.UserName
 Connect-SPOService -Url $SPOAdminSiteUrl -Credential $PSCred   
 Get-SPOSite | % { Write-Output ("Site:" + $_.Url) }  
}