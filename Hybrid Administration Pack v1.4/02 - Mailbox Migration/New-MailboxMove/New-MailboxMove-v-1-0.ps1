#################################################################################################################
###                                                                                                           ###
###  	Script by Terry Munro -                                                                               ###
###     Technical Blog -               http://365admin.com.au                                                 ###
###     Webpage -                      https://www.linkedin.com/in/terry-munro/                               ###
###     TechNet Gallery Scripts -      http://tinyurl.com/TerryMunroTechNet                                   ###
###                                                                                                           ###
###     TechNet Download link -        https://gallery.technet.microsoft.com/Hybrid-Office-365-9b4570a5       ###
###                                                                                                           ###
###     Version 1.0 - 25/06/2017                                                                              ###
###                                                                                                           ###
###     Revision -                                                                                            ###
###               v1.0 - Initial script                                                                       ###
###                                                                                                           ###
#################################################################################################################


####  Notes for Usage  ######################################################################
#                                                                                           #
#  You MUST be using my Hybrid Connection Script for this mailbox creation script to work   #
#                                                                                           #
#  Download my Hybrid Connection Script -                                                   #
#  --- https://gallery.technet.microsoft.com/Office-365-Hybrid-Azure-354dc04c ---           # 
#                                                                                           #
#  Support Guides -                                                                         #
#   - Pre-Requisites - Configuring your PC for Hybrid Admin                                 #
#   - - -  http://www.365admin.com.au/2017/05/how-to-configure-your-desktop-pc-for.html     #      
#   - Usage Guide - Editing the Hybrid connection script                                    # 
#   - - - http://www.365admin.com.au/2017/05/how-to-connect-to-hybrid-exchange.html         #
#                                                                                           #
#   - Editing and using this mailbox creation script                                        #
#   - - - http://www.365admin.com.au/2017/06/hybrid-management-part-07-moving-bulk.html     #
#                                                                                           #
#############################################################################################

$CSVLocation = "C:\Scripts\MailboxMove.csv"

$LocalMailDelivery = "mail.domain.com"

$Tenant = "tenant"


### Step 1 - Assign licenses to user mailboxes

Import-CSV $CSVLocation | foreach {

if ($_.AccountSkuId -ne "") {
	Set-MsolUser -UserPrincipalName $_.UPN -UsageLocation $_.UsageLocation

    Set-MsolUserLicense -UserPrincipalName $_.UPN -AddLicenses $_.AccountSkuId
}

}


### Step 2 - Migrate mailboxes from Exchange Local to Exchange Online

 
Import-CSV $CSVLocation | ForEach-Object { 
New-EXOMoveRequest –identity $_.UPN -BatchName $_.UPN -Remote -RemoteHostName $LocalMailDelivery -TargetDeliveryDomain "$($Tenant).mail.onmicrosoft.com" -RemoteCredential $LocalCredential
}


Write-Host "Mailboxes are now migrating. Please be patient while they move"