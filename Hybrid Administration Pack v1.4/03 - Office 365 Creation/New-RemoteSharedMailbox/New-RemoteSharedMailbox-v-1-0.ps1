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
#   - - - http://www.365admin.com.au/2017/06/hybrid-management-part-09-creating.html        #
#                                                                                           #
#############################################################################################

$CSVLocation = "C:\Scripts\RemoteSharedMailboxes.csv"

$LocalMailDelivery = "mail.domain.com"

$Tenant = "tenant"

$Delay = "600"

### Step 1 - Local Shared Mailbox Creation

Import-CSV $CSVLocation | ForEach-Object {
New-EXLMailbox -Name $_.Name -FirstName $_.FirstName -Initials $_.Initials -Lastname $_.LastName -UserPrincipalName $_.UPN -OrganizationalUnit $_.OU -Password (ConvertTo-SecureString $_.password -AsPlainText -Force) -ResetPasswordOnNextLogon $false 

if ($_.City -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -City $_.City
}


if ($_.Company -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -Company $_.Company
}


if ($_.Department -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -Department $_.Department
}


if ($_.HomePhone -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -HomePhone $_.HomePhone
}


if ($_.MobilePhone -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -MobilePhone $_.MobilePhone
}


if ($_.OfficePhone -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -OfficePhone $_.OfficePhone
}


if ($_.Office -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -Office $_.Office
}


if ($_.PostalCode -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -PostalCode $_.PostalCode
}


if ($_.StreetAddress -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -StreetAddress $_.StreetAddress
}


if ($_.State -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -State $_.State
}


if ($_.Country -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -Country $_.Country
}


if ($_.Title -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -Title $_.Title
}



if ($_.HomePage -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -HomePage $_.HomePage
}


if ($_.Fax -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -Fax $_.Fax
}

           
Get-EXLMailbox -Identity $_.Alias | Set-EXLMailbox  -CustomAttribute1 $_.CustomAttribute1
}


### Step 2 - Sync new local mailboxes to Office 365 via Azure AD Connect

Write-Host "Please wait while I add delegates with Full Access, Send As, and SendOnBehalf permissions"
Start-Sleep -s 120


Import-CSV $CSVLocation | ForEach-Object {

if ($_.SendAs -ne "") {
	Get-EXLMailbox $_.Alias | Add-EXLADPermission -User $_.SendAs -AccessRights 'ExtendedRight' -ExtendedRights 'send as'
}


if ($_.SendOnBehalf -ne "") {
	Get-EXLMailbox $_.Alias | Set-EXLMailbox -GrantSendOnBehalfTo $_.SendOnBehalf
}


if ($_.FullAccess -ne "") {
	Get-EXLMailbox $_.Alias | Add-EXLMailboxpermission -user $_.FullAccess -AccessRights FullAccess -InheritanceType All
}

}



### Step 3 - Sync new local mailboxes to Office 365 via Azure AD Connect
Start-ADSyncSyncCycle -PolicyType Delta

Write-Host "Local Mailboxes created and synched to Office 365"

Start-Sleep -s 300

Write-Host "Now we will migrate the mailboxes to Office 365 because Microsoft don't have a better way"


### Step 4 - Migrating the local User mailboxes to Office 365 because Microsoft don't have a better way

 
Import-CSV $CSVLocation | ForEach-Object { 
New-EXOMoveRequest –identity $_.UPN -BatchName $_.Alias -Remote -RemoteHostName $LocalMailDelivery -TargetDeliveryDomain "$($Tenant).mail.onmicrosoft.com" -RemoteCredential $LocalCredential
}


Write-Host "Mailboxes are now migrating. Please be patient while they move"

Start-Sleep -s $Delay

Write-Host "Please wait while I change the mailbox type to Shared"

Import-CSV $CSVLocation | ForEach-Object { 
Get-EXOMailbox $_.Alias | Set-EXOMailbox -Type Shared
}


