Connect-MsolService
Get-MsolUser -maxresults 9000| Where-Object { $_.isLicensed -eq "TRUE" } | Export-Csv c:\users\scottcroucher\LicensedUsers2.csv 