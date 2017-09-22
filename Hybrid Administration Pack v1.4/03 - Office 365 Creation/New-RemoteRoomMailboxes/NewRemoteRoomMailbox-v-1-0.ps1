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
#   - - - http://www.365admin.com.au/2017/06/hybrid-management-part-10-creating.html        #
#                                                                                           #
#############################################################################################

$CSVLocation = "C:\Scripts\RemoteRoomMailboxes.csv"

### Step 1 - User Creation

Import-CSV $CSVLocation | ForEach-Object {
New-EXLRemoteMailbox -Room -Name $_.Name -FirstName $_.FirstName -Initials $_.Initials -Lastname $_.LastName -UserPrincipalName $_.UPN -OnPremisesOrganizationalUnit $_.OU

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


if ($_.Manager -ne "") {
	Get-ADUser $_.Alias | Set-ADUser -Manager $_.Manager
}


if ($_.CustomAttribute1 -ne "") {
	Get-EXLRemoteMailbox -Identity $_.Alias | Set-EXLRemoteMailbox -CustomAttribute1 $_.CustomAttribute1
}


if ($_.SendOnBehalf -ne "") {
	Get-EXLRemoteMailbox $_.Alias | Set-EXLRemoteMailbox -GrantSendOnBehalfTo $_.SendOnBehalf
}

}



### Step 3 - Sync new local mailboxes to Office 365 via Azure AD Connect
Start-ADSyncSyncCycle -PolicyType Delta


Write-Host "Please wait while the accounts are created in Office 365"
Start-sleep -s 300

Write-Host "Please wait while delegate permissions are applied"


### Step 4 - Assign permissions to the Remote Room Mailbox


Import-CSV $CSVLocation | ForEach-Object {

if ($_.FullAccess -ne "") {
	Add-EXOMailboxpermission -identity $_.UPN -user $_.FullAccess -AccessRights FullAccess -InheritanceType All
}

if ($_.SendAs -ne "") {
	Add-EXORecipientPermission $_.Alias -AccessRights SendAs -Trustee $_.SendAs -Confirm:$false
}

if ($_.ResourceCapacity -ne "") {
	Get-EXOMailbox $_.Alias | Set-EXOMailbox -ResourceCapacity $_.ResourceCapacity
}


if ($_.BookingDelegate -eq "") {
	Get-EXOMailbox -Identity $_.Alias | Set-EXOCalendarProcessing -AllBookInPolicy:$true -AutomateProcessing 'AutoAccept' -AllRequestInPolicy:$false
}


if ($_.BookingDelegate -ne "") {
	Get-EXOMailbox -Identity $_.Alias | Set-EXOCalendarProcessing -AllRequestInPolicy:$true -ResourceDelegates $_.BookingDelegate -AllBookInPolicy:$false -AutomateProcessing 'AutoAccept'
}


}


Write-Host "Remote Room Mailboxes created and synched to Office 365."