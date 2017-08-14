#Import Active Directory Module in PowerShell
Import-Module ActiveDirectory
Import-Csv "D:\O365.csv" | ForEach-Object {
$upn = $_."UPN"
$user = $upn.IndexOf("@")
$left = $upn.Substring(0,$user)
$upn1 = "$left" + "@tenantname.onmicrosoft.com"
Restore-MsolUser -UserPrincipalName $upn
Set-MsolUserPrincipalName -UserPrincipalName $upn -NewUserPrincipalName $upn1
Set-Msoluser -UserPrincipalName $upn1 -ImmutableID "$Null"
$id=(Get-ADUser -Filter {UserPrincipalName -like $upn } -Properties ObjectGUID | select ObjectGUID | foreach {[system.convert]::ToBase64String(([GUID]($_.ObjectGUID)).tobytearray())})
    Set-MSOLUser –UserPrincipalName $upn1 –ImmutableID $id
    Set-MsolUserPrincipalName -UserPrincipalName $upn1 -NewUserPrincipalName $upn
    }