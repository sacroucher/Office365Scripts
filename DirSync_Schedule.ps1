<#
. AUTHOR
    DirSync_Schedule.ps1 v1.3
    Copyright:     Free to use, please leave this header intact  
    Author:        Scott Croucher 
    Company:       Carnnell LTD (www.carnnell.co.uk) 
    Purpose:       The script checks the dir sync schedule when it was last run and when it will run again.

                  
#>

$pshost = Get-Host              # Get the PowerShell Host. 
$pswindow = $pshost.UI.RawUI    # Get the PowerShell Host's UI. 
 
$newsize = $pswindow.windowsize 
$newsize.height = 6             # Set modal window height 
$newsize.width = 55             # Set modal window width  
$pswindow.windowsize = $newsize 
 
$signature = @’ 
[DllImport("user32.dll")] 
public static extern bool SetWindowPos( 
    IntPtr hWnd, 
    IntPtr hWndInsertAfter, 
    int X, 
    int Y, 
    int cx, 
    int cy, 
    uint uFlags); 
‘@ 
  
$type = Add-Type -MemberDefinition $signature -Name SetWindowPosition -Namespace SetWindowPos -Using System.Text -PassThru 
 
$handle = (Get-Process -id $Global:PID).MainWindowHandle 
$alwaysOnTop = New-Object -TypeName System.IntPtr -ArgumentList (-1) 
$type::SetWindowPos($handle, $alwaysOnTop, 0, 0, 0, 0, 0x0003) 
 
Import-Module MSonline                    #Establish session to cloud 
$O365Cred=Get-Credential 
$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $O365Cred -Authentication Basic -AllowRedirection 
Import-PSSession $O365Session -AllowClobber 
Set-ExecutionPolicy Unrestricted 
Connect-MsolService -Credential $O365Cred 
#Import-Module LyncOnlineConnector 
Clear-Host 
 
 
While($true)                              #Looping construct for DirSync Update 
{ 
    Clear-Host 
 
    $t = Get-Date 
 
    Write-Host "DIR SYNC TIMER - Last Run at: " $t  -ForegroundColor Red `r`n 
     
    Start-Sleep -Seconds 2 
 
    $x = Get-MsolCompanyInformation | select -ExpandProperty LastDirSyncTime 
 
    $y = New-TimeSpan -Hours 2 
    $y2 = New-TimeSpan + Hours 1
 
    $z = ($x) - $y  
    $z2 = ($z) + $y2 
 
     
    Write-Host "Last DirSync occurred at: " $z -ForegroundColor Cyan `r`n 
    Write-Host "Next DirSync will occur at: "  $z2 -ForegroundColor Cyan 
 
 
    Start-Sleep -Seconds 600 
     
     
       
} 