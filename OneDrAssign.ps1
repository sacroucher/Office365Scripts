<#
. AUTHOR
    OneDriveMapper v2.23 
    Copyright:     Free to use, please leave this header intact  
    Author:        Scott Croucher 
    Company:       Carnnell LTD (www.carnnell.co.uk) 
    Purpose:       This script maps your Office 365 OneDrive folder automatically upon logon, if not already present

.DESCRIPTION
    This script maps your Office 365 OneDrive folder automatically upon logon, if not already present
.PARAMETER
    Modify the example URL and check if something can be done about URL encoding of SP url's to prevent file locking  
    Add detection of zone specific GPO enforcement of protected mode (which overrides the autoProtectedMode setting) 
.PARAMETER  
    To use ADFS SSO, your federation URL (fs.domain.com) should be in Internet Explorer's Intranet Sites, windows authentication should be enabled in IE.  
    Users should also have the same login UPN on their client, as in Office 365. The script will attempt to log into Office 365 using either the full   
    user UPN from Active Directory (if lookupUPNbySAM is enabled) or by using the SamAccountName + the domain name.  
    If you run any type of mapping scripts to your drive, make sure this script runs first or the drive will not be available.  
    If you use a desktop management tool like RES PowerFuse, make sure your users are allowed to start a COM object and run powershell scripts.  
#>  
  
 

 
########  
#Configuration  
########  
$domain             = "OGD.NL"                    #This should be your domain name in O365, and your UPN in Active Directory, for example: ogd.nl  
$driveLetter         = "O:"                         #This is the driveletter you'd like to use for OneDrive, for example: Z:  
$driveLabel         = "OGD"                     #If you enter a name here, the script will attempt to label the drive with this value  
$O365CustomerName    = "ogd"                        #This should be the name of your tenant (example, ogd as in ogd.onmicrosoft.com)  
$logfile            = ($env:APPDATA + "\OneDriveMapper.log")    #Logfile in case of errors  
$debugmode          = $False                      #Set to $True for debugging purposes. You'll be able to see the script navigate in Internet Explorer  
$lookupUPNbySAM     = $True                     #Look up the user's UPN by the SAMAccountName, use this if your UPN doesn't match your SamAccountName or if you have multiple domains  
$forceUserName      = ""                        #if anything is entered here, there will be no UPN lookup and the domain will be ignored. This is useful for machines that aren't domain joined.  
$forcePassword      = ""                        #if anything is entered here, the user won't be prompted for a password. This function is not recommended, as your password could be stolen from this file  
$autoProvision      = $True                     #If a user has never accessed his/her OneDrive, it won't be ready yet, this setting attempts to log in once to provision it before mapping  
$autoProtectedMode  = $True                     #Automatically temporarily disable IE Protected Mode if it is enabled. ProtectedMode has to be disabled for the script to function  
$provisionWaitTime  = 30                        #Time to wait in seconds for Office 365 to provision a personal sharepoint site (MySite) or to abort  
$adfsWaitTime       = 10                         #Amount of seconds to allow for ADFS redirects, if set too low, the script may fail while just waiting for a slow ADFS redirect, this is because the IE object will report being ready even though it is not.  Set to 0 if not using ADFS.  
$libraryName        = "Documents"               #leave this default, unless you wish to map a non-default library you've created  
$sharepointURL      = ""                        #Write the full Sharepoint Library URL here if you want to map to a Sharepoint shared library instead of your own personal folder  
#Example sharepoint URL: https://lieben.sharepoint.com/TeamSite/Documents or https://lieben.sharepoint.com/TeamSite/Gedeelde documenten  
$restart_explorer   = $False                    #Set to true if drives are always invisible after the script runs, this will restart explorer.exe after mapping the drive  
$autoKillIE         = $True                     #Kill any running Internet Explorer processes prior to running the script to prevent security errors when mapping  
$abortIfNoAdfs      = $False                    #If set to True, will stop the script if no ADFS server has been detected during login 
$buttonText         = "Login"                   #Text of the button on the password input popup box 
$adfsLoginInput     = "userNameInput"           #Only modify this if you have a customized (skinned) ADFS implementation 
$adfsPwdInput       = "passwordInput" 
$adfsButton         = "submitButton" 
  
########  
#Required resources  
########  
$mapresult = $False  
$protectedModeValues = @{}  
$privateSuffix = "-my"  
  
ac $logfile "-----$(Get-Date) OneDriveMapper V2.23 by $($env:USERNAME) on $($env:COMPUTERNAME) Session log-----"  
  
#remove the -my suffix when we're mapping to a Sharepoint Library  
if($sharepointURL.Length -gt 0) {  
    $privateSuffix = ""  
}  
  
Write-Host "One moment please, your $($driveLetter) drive is connecting..."  
  
$domain = $domain.ToLower()  
$O365CustomerName = $O365CustomerName.ToLower()  
$forceUserName = $forceUserName.ToLower()  
 
#region basicFunctions 
function lookupUPN{  
    try{  
        $objDomain = New-Object System.DirectoryServices.DirectoryEntry  
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher  
        $objSearcher.SearchRoot = $objDomain  
        $objSearcher.Filter = �(&(objectCategory=User)(SAMAccountName=$Env:USERNAME))�  
        $objSearcher.SearchScope = �Subtree�  
        $objSearcher.PropertiesToLoad.Add(�userprincipalname�) | Out-Null  
        $results = $objSearcher.FindAll()  
        return $results[0].Properties.userprincipalname  
    }catch{  
        ac $logfile "Failed to lookup username, active directory connection failed, please disable lookupUPN"  
        abort_OM  
    }  
}  
  
function CustomInputBox([string] $title, [string] $message)   
{  
    if($forcePassword.Length -gt 2) {  
        return $forcePassword  
    }  
    $objBalloon = New-Object System.Windows.Forms.NotifyIcon   
    $objBalloon.BalloonTipIcon = "Info"  
    $objBalloon.BalloonTipTitle = "OneDriveMapper"   
    $objBalloon.BalloonTipText = "OneDriveMapper - www.carnnell.co.uk "  
    $objBalloon.Visible = $True   
    $objBalloon.ShowBalloonTip(10000)  
  
    $userForm = New-Object 'System.Windows.Forms.Form'  
    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'  
    $Form_StateCorrection_Load=  
    {  
        $userForm.WindowState = $InitialFormWindowState  
    }  
  
    $userForm.Text = "$title"  
    $userForm.Size = New-Object System.Drawing.Size(350,200)  
    $userForm.StartPosition = "CenterScreen"  
    $userForm.AutoSize = $False  
    $userForm.MinimizeBox = $False  
    $userForm.MaximizeBox = $False  
    $userForm.SizeGripStyle= "Hide"  
    $userForm.WindowState = "Normal"  
    $userForm.FormBorderStyle="Fixed3D"  
    $userForm.KeyPreview = $True  
    $userForm.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$userForm.Close()}})    
    $OKButton = New-Object System.Windows.Forms.Button  
    $OKButton.Location = New-Object System.Drawing.Size(105,110)  
    $OKButton.Size = New-Object System.Drawing.Size(95,23)  
    $OKButton.Text = $buttonText  
    $OKButton.Add_Click({$userForm.Close()})  
    $userForm.Controls.Add($OKButton)  
    $userLabel = New-Object System.Windows.Forms.Label  
    $userLabel.Location = New-Object System.Drawing.Size(10,20)  
    $userLabel.Size = New-Object System.Drawing.Size(300,50)  
    $userLabel.Text = "$message"  
    $userForm.Controls.Add($userLabel)   
    $objTextBox = New-Object System.Windows.Forms.TextBox  
    $objTextBox.UseSystemPasswordChar = $True  
    $objTextBox.Location = New-Object System.Drawing.Size(70,75)  
    $objTextBox.Size = New-Object System.Drawing.Size(180,20)  
    $userForm.Controls.Add($objTextBox)   
    $userForm.Topmost = $True  
    $userForm.TopLevel = $True  
    $userForm.ShowIcon = $True  
    $userForm.Add_Shown({$userForm.Activate();$objTextBox.focus()})  
    $InitialFormWindowState = $userForm.WindowState  
    $userForm.add_Load($Form_StateCorrection_Load)  
    [void] $userForm.ShowDialog()  
    return $objTextBox.Text  
}  
  
function labelDrive{  
    Param(  
    [String]$lD_DriveLetter,  
    [String]$lD_MapURL,  
    [String]$lD_DriveLabel  
    )  
  
    #try to set the drive label  
    if($lD_DriveLabel.Length -gt 0){  
        ac $logfile "A drive label has been specified, attempting to set the label for $($lD_DriveLetter)"  
        try{  
            $regURL = $lD_MapURL.Replace("\","#")  
            $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($regURL)"  
            New-Item -Path $path -Value "default value" �Force  
            New-ItemProperty -Path $path -Name "_CommentFromDesktopINI"  
            New-ItemProperty -Path $path -Name "_LabelFromDesktopINI"  
            New-ItemProperty -Path $path -Name "_LabelFromReg" -Value $lD_DriveLabel  
            $regURL = $regURL.Replace("DavWWWRoot#","")  
            $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($regURL)"  
            New-Item -Path $path -Value "default value" �Force  
            New-ItemProperty -Path $path -Name "_CommentFromDesktopINI"  
            New-ItemProperty -Path $path -Name "_LabelFromDesktopINI"  
            New-ItemProperty -Path $path -Name "_LabelFromReg" -Value $lD_DriveLabel  
            ac $logfile "Label has been set to $($lD_DriveLabel)"  
  
        }catch{  
            ac $logfile "Failed to set the drive label registry keys"  
            ac $logfile $error[0]  
        }  
  
    }  
}  
function MapDrive{  
    Param(  
    [String]$MD_DriveLetter,  
    [String]$MD_MapURL,  
    [String]$MD_DriveLabel  
    )  
    ac $logfile "Mapping target: $($MD_MapURL)`n"  
    $del = NET USE $MD_DriveLetter /DELETE 2>&1  
    $out = NET USE $MD_DriveLetter $MD_MapURL /PERSISTENT:YES 2>&1  
    if($LASTEXITCODE -ne 0){  
        if((Get-Service -Name WebClient).Status -ne "Running"){  
            ac $logfile "CRITICAL ERROR: OneDriveMapper detected that the WebClient service was not started, please ensure this service is always running!`n"  
        }  
        ac $logfile "Failed to map $($MD_DriveLetter) to $($MD_MapURL), error: $($LASTEXITCODE) $($out)`n"  
        return $False  
    }  
    if([System.IO.Directory]::Exists($MD_DriveLetter)){  
        #set drive label  
        labelDrive $MD_DriveLetter $MD_MapURL $MD_DriveLabel  
        ac $logfile "$($MD_DriveLetter) mapped successfully`n"  
        if($restart_explorer){  
            ac $logfile "Restarting Explorer.exe to make the drive visible"  
            #kill all running explorer instances of this user  
            $explorerStatus = Get-ProcessWithOwner explorer  
            if($explorerStatus -eq 0){  
                ac $logfile "WARNING: no instances of Explorer running yet, at least one should be running"  
            }elseif($explorerStatus -eq -1){  
                ac $logfile "ERROR Checking status of Explorer.exe: unable to query WMI"  
            }else{  
                ac $logfile "Detected running Explorer processes, attempting to shut them down..."  
                foreach($Process in $explorerStatus){  
                    try{  
                        Stop-Process $Process.handle | Out-Null  
                        ac $logfile "Stopped process with handle $($Process.handle)"  
                    }catch{  
                        ac $logfile "Failed to kill process with handle $($Process.handle)"  
                    }  
                }  
            }  
        }  
        return $True  
    }else{  
        if($LASTEXITCODE -eq 0){  
            ac $logfile "failed to contact $($MD_DriveLetter) after mapping it to $($MD_MapURL), check if the URL is valid"  
            ac $logfile $error[0]  
            #($error[0]|format-list -force)  
        }  
        return $False  
    }  
}  
  
function revertProtectedMode(){  
    ac $logfile "autoProtectedMode is set to True, reverting to old settings"  
    try{  
        for($i=1; $i -lt 4; $i++){  
            if($protectedModeValues[$i] -ne $Null){  
                ac $logfile "Setting zone $i back to $($protectedModeValues[$i])"  
                Set-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500"  -Value $protectedModeValues[$i] -ErrorAction SilentlyContinue  
            }  
        }  
    }  
    catch{  
        ac $logfile "Failed to modify registry keys to change ProtectedMode back to the original settings"  
        ac $logfile $error[0]  
    }  
}  
 
function abort_OM{  
    #find and kill all active COM objects for IE 
    $ie.Quit() 
    $shellapp = New-Object -ComObject "Shell.Application" 
    $ShellWindows = $shellapp.Windows() 
    for ($i = 0; $i -lt $ShellWindows.Count; $i++) 
    { 
      if ($ShellWindows.Item($i).FullName -like "*iexplore.exe") 
      {
        $del = $ShellWindows.Item($i) 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($del)  
      }
    } 
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shellapp)  
    if($autoProtectedMode){  
        revertProtectedMode  
    }  
    Exit  
}  
  
function askForPassword{  
    do{  
        $askAttempts++  
        ac $logfile "asking user for password`n"  
        try{  
            $password = CustomInputBox "Microsoft Office 365 OneDrive" "Please enter the password for $($userUPN.ToLower()) to access $($driveLetter)"  
        }catch{  
            ac $logfile "failed to display a password input box, exiting`n"  
            abort_OM               
        }  
    }  
    until($password.Length -gt 0 -or $askAttempts -gt 2)  
    if($askAttempts -gt 3) {  
        ac $logfile "user refused to enter a password, exiting`n"  
        abort_OM  
    }else{  
        return $password  
    }  
}  
  
function Get-ProcessWithOwner {  
    param(  
        [parameter(mandatory=$true,position=0)]$ProcessName  
    )  
    $ComputerName=$env:COMPUTERNAME  
    $UserName=$env:USERNAME  
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($(New-Object System.Management.Automation.PSPropertySet(�DefaultDisplayPropertySet�,[string[]]$('ProcessName','UserName','Domain','ComputerName','handle'))))  
    try {  
        $Processes = Get-wmiobject -Class Win32_Process -ComputerName $ComputerName -Filter "name LIKE '$ProcessName%'"  
    } catch {  
        return -1
    }  
    if ($Processes -ne $null) {  
        $OwnedProcesses = @()  
        foreach ($Process in $Processes) {  
            if($Process.GetOwner().User -eq $UserName){  
                $Process |   
                Add-Member -MemberType NoteProperty -Name 'Domain' -Value $($Process.getowner().domain)  
                $Process |  
                Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value $ComputerName   
                $Process |  
                Add-Member -MemberType NoteProperty -Name 'UserName' -Value $($Process.GetOwner().User)   
                $Process |   
                Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers  
                $OwnedProcesses += $Process  
            }  
        }  
        return $OwnedProcesses  
    } else {  
        return 0  
    }  
  
}  
#endregion 
 
#region loginFunction 
function login(){  
    ac $logfile "Login attempt at Office 365 signin page"
    #click to open up the login menu  
    do {sleep -m 100} until (-not ($ie.Busy))   
    if($ie.document.GetElementById("_link").tagName -ne $Null){  
       $ie.document.GetElementById("_link").click()   
       ac $logfile "Found sign in elements type 1 on Office 365 login page, proceeding"  
    }elseif($ie.document.GetElementById("use_another_account").tagName -ne $Null){  
       $ie.document.GetElementById("use_another_account").click()  
       ac $logfile "Found sign in elements type 2 on Office 365 login page, proceeding"  
    }elseif($ie.document.GetElementById("use_another_account_link").tagName -ne $Null){  
       $ie.document.GetElementById("use_another_account_link").click()  
       ac $logfile "Found sign in elements type 3 on Office 365 login page, proceeding"  
    }elseif($ie.document.GetElementById("_use_another_account_link").tagName -ne $Null){  
       $ie.document.GetElementById("_use_another_account_link").click()  
       ac $logfile "Found sign in elements type 4 on Office 365 login page, proceeding"  
    }elseif($ie.document.GetElementById("cred_keep_me_signed_in_checkbox").tagName -ne $Null){  
       ac $logfile "Found sign in elements type 5 on Office 365 login page, proceeding"  
    }else{  
       ac $logfile "Script was unable to find browser controls on the login page and cannot continue, please check your safe-sites or verify these elements are present"  
       abort_OM  
    }  
    do {sleep -m 100} until (-not ($ie.Busy))   
  
  
    #attempt to trigger redirect to detect if we're using ADFS automatically  
    try{  
        ac $logfile "attempting to trigger a redirect to ADFS"  
        $ie.document.GetElementById("cred_keep_me_signed_in_checkbox").click()  
        $ie.document.GetElementById("cred_userid_inputtext").value = $userUPN  
        do {sleep -m 100} until (-not ($ie.Busy))   
        $ie.document.GetElementById("cred_password_inputtext").click()  
        do {sleep -m 100} until (-not ($ie.Busy))   
    }catch{  
        ac $logfile "Failed to find the correct controls at $($ie.LocationURL) to log in by script, check your browser and proxy settings or check for an update of this script`n"  
        abort_OM   
    }  
  
    sleep -s 2  
    $redirecting = $True  
    $redirWaited = 0  
    while($redirecting){  
        sleep -m 500  
        try{ 
            $found_Splitter = $ie.document.GetElementById("aad_account_tile_link").tagName 
        }catch{ 
            $found_Splitter = $Null 
        } 
        #Select business account if the option is presented  
        if($found_Splitter -ne $Null){  
            $ie.document.GetElementById("aad_account_tile_link").click()  
            ac $logfile "Login splitter detected, your account is both known as a personal and business account, selecting business account.."  
            sleep -s 2  
            $redirWaited += 2 
        }  
        #check if the COM object is healthy, otherwise we're running into issues  
        if($ie.HWND -eq $null){  
            ac $logfile "ERROR: the browser object was Nulled during login, this means IE ProtectedMode or other security settings are blocking the script."  
            abort_OM  
        }  
 
        #this is the ADFS login control ID, modify this if you have a custom IdP 
        try{ 
            $found_ADFSControl = $ie.document.GetElementById($adfsLoginInput).tagName 
        }catch{ 
            $found_ADFSControl = $Null 
            ac $logfile "ADFS userNameInput element not found: $($Error[0]) with method 1" 
        } 
        #try alternative method for selecting the ID  
        if($found_ADFSControl.Length -lt 1){ 
            try{ 
                $found_ADFSControl = $ie.Document.IHTMLDocument3_getElementById($adfsLoginInput).tagName 
            }catch{ 
                $found_ADFSControl = $Null 
                ac $logfile "ADFS userNameInput element not found: $($Error[0]) with method 2" 
            } 
        } 
        $redirWaited += 0.5  
        #found ADFS control 
        if($found_ADFSControl){ 
            ac $logfile "ADFS Control found, we were redirected to: $($ie.LocationURL)"  
            $redirecting = $False 
            $useADFS = $True 
        }  
 
        if($redirWaited -ge $adfsWaitTime){  
            ac $logfile "waited for more than $adfsWaitTime to get redirected to ADFS, checking if we were properly redirected or attempting normal signin"  
            $useADFS = $False     
            $redirecting = $False    
        }  
    }      
 
    #if not using ADFS, sign in  
    if($useADFS -eq $False){  
        if($abortIfNoAdfs){ 
            ac $logfile "abortIfNoAdfs was set to true, ADFS was not detected, script is exiting" 
            abort_OM 
        } 
        if($ie.LocationURL.StartsWith($baseURL)){ 
            #we've been logged in, we can abort the login function  
            ac $logfile "login detected, login function succeeded, final url: $($ie.LocationURL)"  
            return $True              
        } 
        try{  
            $ie.document.GetElementById("cred_password_inputtext").value = askForPassword  
            $ie.document.GetElementById("cred_sign_in_button").click()  
            do {sleep -m 100} until (-not ($ie.Busy)) 
        }catch{  
            ac $logfile "Failed to find the correct controls at $($ie.LocationURL) to log in by script, check your browser and proxy settings or check for an update of this script`n"  
            abort_OM   
        }  
    }else{  
        #check if logged in now automatically after ADFS redirect  
        if($ie.LocationURL.StartsWith($baseURL)){  
            #we've been logged in, we can abort the login function  
            ac $logfile "login detected, login function succeeded, final url: $($ie.LocationURL)"  
            return $True  
        }  
    }  
  
    #Not logged in automatically, so ADFS requires us to sign in  
    do {sleep -m 100} until (-not ($ie.Busy)) 
  
    #Check if we arrived at a 404, or an actual page  
    if($ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*") {  
        ac $logfile "We received a 404 error after our signin attempt, this script cannot continue"  
        abort_OM           
    }  
  
    #check if logged in now  
    if($ie.LocationURL.StartsWith($baseURL)){  
        #we've been logged in, we can abort the login function  
        ac $logfile "login detected, login function succeeded, final url: $($ie.LocationURL)"  
        return $True  
    }else{  
        if($useADFS){  
            ac $logfile "ADFS did not automatically sign us on, attempting to enter credentials at $($ie.LocationURL)"  
            try{  
                $ie.document.GetElementById($adfsLoginInput).value = $userUPN  
                $ie.document.GetElementById($adfsPwdInput).value = askForPassword  
                $ie.document.GetElementById($adfsButton).click()  
                do {sleep -m 100} until (-not ($ie.Busy))   
                sleep -s 1  
                do {sleep -m 100} until (-not ($ie.Busy))    
            }catch{  
                ac $logfile "Failed to find the correct controls at $($ie.LocationURL) using method 1 to log in by script, will try method 2`n"  
                $tryMethod2 = $True 
            }  
            if($tryMethod2 -eq $True){ 
                try{  
                    $ie.document.IHTMLDocument3_getElementById($adfsLoginInput).value = $userUPN  
                    $ie.document.IHTMLDocument3_getElementById($adfsPwdInput).value = askForPassword  
                    $ie.document.IHTMLDocument3_getElementById($adfsButton).click()  
                    do {sleep -m 100} until (-not ($ie.Busy))   
                    sleep -s 1  
                    do {sleep -m 100} until (-not ($ie.Busy))    
                }catch{  
                    ac $logfile "Failed to find the correct controls at $($ie.LocationURL) using method 2 to log in by script, check your browser and proxy settings or modify this script to match your ADFS form`n"  
                    $tryMethod2 = $True 
                    abort_OM  
                }  
            } 
            do {sleep -m 100} until (-not ($ie.Busy))    
            #check if logged in now  
            if($ie.LocationURL.StartsWith($baseURL)){  
                #we've been logged in, we can abort the login function  
                ac $logfile "login detected, login function succeeded, final url: $($ie.LocationURL)"  
                return $True  
            }else{  
                ac $logfile "We attempted to login with ADFS, but did not end up at the expected location. Detected url: $($ie.LocationURL), expected URL: $($baseURL)"  
                abort_OM  
            }  
        }else{  
            ac $logfile "We attempted to login without using ADFS, but did not end up at the expected location. Detected url: $($ie.LocationURL), expected URL: $($baseURL)"  
            abort_OM  
        }  
    }  
}  
#endregion 
 
 
  
#get user login  
if($lookupUPNbySAM){  
    ac $logfile "lookupUPNbySAM is set to True -> Using UPNlookup by SAMAccountName feature`n"  
    $userUPN = lookupUPN  
}else{  
    $userUPN = ([Environment]::UserName)+"@"+$domain  
    ac $logfile "lookupUPNbySAM is set to False -> Using $userUPN from the currently logged in username + $domain`n"  
}  
if($forceUserName.Length -gt 2){  
    ac $logfile "A username was already specified in the script configuration: $($forceUserName)`n"  
    $userUPN = $forceUserName  
}  
 
#region flightChecks 
#Check if Office 365 libraries have been installed  
try{  
    [System.IO.File]::Exists("$(${env:ProgramFiles(x86)})\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll")  
}catch{  
    ac $logfile "Possible critical error: Microsoft Office installation not detected, script may fail"  
}  
  
  
#Check if Zone Configuration is on a per machine or per user basis, then check the zones  
$zoneFound = $False 
$BaseKeypath = "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"  
try{  
    $IEMO = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "Security HKLM only" -ErrorAction Stop | Select-Object 'Security HKLM only'  
}catch{  
    ac $logfile "NOTICE: $($BaseKeypath)\Security HKLM only not found in registry, your zone configuration could be set on both levels"  
}  
if($IEMO.'Security HKLM only' -eq 1){  
    ac $logfile "NOTICE: $($BaseKeypath)\Security HKLM only found in registry and set to 1, your zone configuration is set on a machine level"     
}else{  
    #Check if sharepoint tenant is in safe sites list of the user  
    $BaseKeypath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)"  
    $BaseKeypath2 = "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)"  
    $zone = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https  
    if($zone -eq $Null){  
        $zone = Get-ItemProperty -Path "$($BaseKeypath2)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https    
    }      
    if($zone.https -eq 2){  
        ac $logfile "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on user level"   
        $zoneFound = $True 
    } 
}  
#Check if sharepoint tenant is in safe sites list of the machine  
$BaseKeypath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)"  
$BaseKeypath2 = "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)"  
$zone = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https  
if($zone -eq $Null){  
    $zone = Get-ItemProperty -Path "$($BaseKeypath2)\" -Name "https" -ErrorAction SilentlyContinue | Select-Object https    
}      
if($zone.https -eq 2){  
    ac $logfile "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on machine level"   
    $zoneFound = $True 
} 
if($zoneFound -eq $False){ 
    ac $logfile "Possible critical error: $($O365CustomerName)$($privateSuffix).sharepoint.com not found in IE Trusted Sites on user or machine level, the script will likely fail" 
} 
  
  
#Check if IE FirstRun is disabled  
$BaseKeypath = "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main"  
try{  
    $IEFR = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "DisableFirstRunCustomize" -ErrorAction Stop | Select-Object DisableFirstRunCustomize  
}catch{  
    ac $logfile "WARNING: $($BaseKeypath)\DisableFirstRunCustomize not found in registry, if script hangs this may be due to the First Run popup in IE"  
}  
if($IEFR.DisableFirstRunCustomize -ne 1){  
    ac $logfile "Possible error: $($BaseKeypath)\DisableFirstRunCustomize not set"     
}  
  
  
#Check if WebDav file locking is enabled  
$BaseKeypath = "HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\"  
try{  
    $wdlocking = Get-ItemProperty -Path "$($BaseKeypath)\" -Name "SupportLocking" -ErrorAction Stop | Select-Object SupportLocking  
}catch{  
    ac $logfile "WARNING: HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters registry location not accessible"  
}  
if($wdlocking.SupportLocking -ne 0){  
    ac $logfile "WARNING: WebDav File Locking support is enabled, this could cause files to become locked in your OneDrive"     
}  
#endregion 
  
#translate to URLs  
$userURL = ($userUPN.replace(".","_")).replace("@","_").ToLower()  
$mapURL = ("\\"+$O365CustomerName+$privateSuffix+".sharepoint.com@SSL\DavWWWRoot\personal\"+$userURL+"\"+$libraryName)  
$mapURLpersonal = ("\\"+$O365CustomerName+"-my.sharepoint.com@SSL\DavWWWRoot\personal\")  
$baseURL = ("https://"+$O365CustomerName+$privateSuffix+".sharepoint.com")  
  
#if we're just mapping to sharepoint, the target URL's need to be rewritten  
if($sharepointURL.Length -gt 0){  
    $mapURL = $sharepointURL.Replace("https://","\\")  
    $mapURL = $mapURL.Replace("sharepoint.com/","sharepoint.com@SSL\")  
    $mapURL = $mapURL.Replace("/","\")  
}  
  
#check if drivemap already exists 
if([System.IO.Directory]::Exists($driveLetter)){  
    $reMap = $False 
    if($sharepointURL.Length -gt 0){  
        #check if mapped path is to the sharepoint location 
        if((Get-PSDrive $driveLetter).DisplayRoot -eq $mapURL){ 
            ac $logfile "the mapped url for $driveLetter matches the desired Sharepoint URL of $mapURL, no need to remap" 
        }else{ 
            ac $logfile "the mapped url for $driveLetter does not match the desired Sharepoint URL of $mapURL" 
            $reMap = $True 
        } 
    }else{ 
        #check if mapped path is to at least the personal folder on Onedrive for Business, username detection would require a full login and slow things down 
        if((Get-PSDrive $driveLetter).DisplayRoot.StartsWith($mapURLpersonal)){ 
            ac $logfile "the mapped url for $driveLetter matches the expected partial URL of $mapURLpersonal, no need to remap" 
        }else{ 
            ac $logfile "the mapped url for $driveLetter does not match the expected partial URL of $mapURLpersonal" 
            $reMap = $True 
        }     
    } 
    #Set the drivelabel 
    labelDrive $driveLetter $mapURL $driveLabel  
 
    #act on previous remap check 
    if($reMap){ 
        #First delete the mapping 
        $out = NET USE $driveLetter /DELETE 2>&1  
        if($LASTEXITCODE -ne 0){ 
            ac $logfile "Failed to unmap $driveLetter : $LASTEXITCODE $out, exiting" 
            abort_OM  
        }else{ 
            ac $logfile "Unmapped $driveLetter" 
        } 
    }else{ 
        abort_OM  
    } 
          
    
}  
  
ac $logfile "Driveletter $($driveLetter) not found -> script will attempt to map $($driveLetter)"  
  
#load windows libraries to display things to the user  
try{  
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")  
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")   
}catch{  
    ac $logfile "Error loading windows forms libraries, script will not be able to display a password input box"  
}  
  
ac $logfile "Base URL: $($baseURL) `n"  
 
#Start IE and stop it once to make sure IE sets default registry keys  
if($autoKillIE){  
    #start invisible IE instance  
    $global:ie = new-object -com InternetExplorer.Application  
    $ie.visible = $debugmode  
    sleep 2  
  
    #kill all running IE instances of this user  
    $ieStatus = Get-ProcessWithOwner iexplore  
    if($ieStatus -eq 0){  
        ac $logfile "WARNING: no instances of Internet Explorer running yet, at least one should be running"  
    }elseif($ieStatus -eq -1){  
        ac $logfile "ERROR Checking status of iexplore.exe: unable to query WMI"  
    }else{  
        ac $logfile "autoKillIE enabled, stopping IE processes"  
        foreach($Process in $ieStatus){  
                Stop-Process $Process.handle -ErrorAction SilentlyContinue 
                ac $logfile "Stopped process with handle $($Process.handle)" 
        }  
    }  
}else{  
    ac $logfile "autoKillIE disabled, IE processes not stopped. This may cause the script to fail for users with a clean/new profile"  
}  
 
if($autoProtectedMode){  
    ac $logfile "autoProtectedMode is set to True, disabling ProtectedMode temporarily"  
    $BaseKeypath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\"  
      
    #store old values and change new ones  
    try{  
        for($i=1; $i -lt 4; $i++){  
            $curr = Get-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500" -ErrorAction SilentlyContinue | select -exp 2500  
            if($curr -ne $Null){  
                $protectedModeValues[$i] = $curr  
                ac $logfile "Zone $i was set to $curr, setting it to 3"  
            }else{ 
                $protectedModeValues[$i] = 0  
                ac $logfile "Zone $i was not yet set, setting it to 3"  
            } 
            Set-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500"  -Value "3" -ErrorAction Stop 
        }  
    }  
    catch{  
        ac $logfile "Failed to modify registry keys to autodisable ProtectedMode $($error[0])"  
    }  
}else{ 
    ac $logfile "autoProtectedMode is set to False, IE ProtectedMode will not be disabled temporarily" 
} 
  
#start invisible IE instance  
try{  
    $global:ie = new-object -com InternetExplorer.Application -ErrorAction Stop 
    $ie.visible = $debugmode  
}catch{  
    ac $logfile "failed to start Internet Explorer COM Object, check user permissions or already running instances`n$($error[0])"   
    abort_OM  
}  
 
#navigate to the base URL of the tenant's Sharepoint to check if it exists  
try{  
    $ie.navigate($baseURL)  
    do {sleep -m 100} until (-not ($ie.Busy))   
}catch{  
    ac $logfile "Failed to browse to the Office 365 Sign in page, this is a fatal error $($error[0])`n"  
    abort_OM  
}  
  
#check if we got a 404 not found  
if($ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*") {  
    ac $logfile "Failed to browse to the Office 365 Sign in page, did you set up O365CustomerName correctly? If it is new, this may take a while.`n"  
    ac $logfile "Debug: attempted url: $($baseURL), exiting script"  
    abort_OM  
}  
  
#check if the COM object is healthy, otherwise we're running into issues  
if($ie.HWND -eq $null){  
    ac $logfile "ERROR: attempt to navigate to $($baseURL) caused the IE scripting object to be nulled. This means your security settings are too high (1)."  
    abort_OM  
}  
ac $logfile "current URL: $($ie.LocationURL)"  
  
#log in  
if($ie.LocationURL.StartsWith($baseURL)){  
    ac $logfile "You were already logged in, skipping login attempt, please note this may fail if you did not log in with a persistent cookie"  
}else{  
    #Check and log if Explorer is running  
    $explorerStatus = Get-ProcessWithOwner explorer  
    if($explorerStatus -eq 0){  
        ac $logfile "WARNING: no instances of Explorer running yet, expected at least one running"  
    }elseif($explorerStatus -eq -1){  
        ac $logfile "ERROR Checking status of explorer.exe: unable to query WMI"  
    }else{  
        ac $logfile "Detected running explorer process"  
    }  
    login  
}  
  
#check autoprovisioning, only if no sharepointURL has been defined  
if($ie.locationURL.Contains("MyBraryFirstRun.aspx") -and $sharepointURL.Length -lt 1){  
    #We are using autoprovision  
    if($autoProvision){  
        ac $logfile "autoProvision is set and we were redirected to the FirstRun page`n"  
        $timeSpent = 1  
        #attempt to visit the onedrive URL until timeout  
        do{  
            $ie.navigate($baseURL)  
            do {sleep -m 100} until (-not ($ie.Busy))   
            do {sleep -m 100} until ($ie.ReadyState -eq 4 -or $ie.ReadyState -eq 0)   
            if($ie.LocationURL.StartsWith($baseURL+"/personal/") -and $ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*" -eq $False){  
                $timeSpent = $provisionWaitTime +1  
                ac $logfile "OneDrive provisioning has completed"  
            }else{  
                ac $logfile "OneDrive still being provisioned, we have waited $($timeSpent) seconds..."  
                Sleep -s 5  
            }  
            $timeSpent += 5  
        }until($timeSpent -gt $provisionWaitTime)  
        #Exit if we didn't end up at the personal URL  
        if($ie.LocationURL.StartsWith($baseURL+"/personal/") -ne $True){  
            ac $logfile "OneDrive still being provisioned, we have waited the maximum allowed time of $($provisionWaitTime) seconds, aborting drivemapping."  
            abort_OM  
        }  
    #else, if firstrun   
    }else{  
        ac $logfile "We were redirected to the FirstRun page of your OneDrive, exiting because autoProvision is set to false. Consider using autoProvision. Exiting script.`n"  
        abort_OM  
    }  
}  
  
#try to determine the username from the URL after we've logged in, only if mapping to OneDrive  
if($sharepointURL.Length -lt 1){  
    $user_detected = $False  
    $detected_loops = 0  
    do{  
        $detected_loops += 1  
        if($detected_loops -gt 10){  
            ac $logfile "Failed to get the username, aborting"  
            abort_OM  
        }  
        $url = $ie.LocationURL  
        $start = $url.IndexOf("/personal/")+10  
        if($start -le 10 -or $url.IndexOf("mysiteredirect.aspx") -gt 0){  
            ac $logfile "Failed to find /personal/ in the URL: $url, waiting 5 seconds and trying again"  
            sleep 5  
        }else{  
            $end = $url.IndexOf("/",$start)  
            try{  
                $userURL = $url.Substring($start,$end-$start)  
                $mapURL = $mapURLpersonal + $userURL + "\" + $libraryName  
            }catch{  
                ac $logfile "WARNING: could not get username from URL, defaulting to user login name`n"   
            }  
            $browseURL = $baseURL + "/personal/" + $userURL + "/" + $libraryName  
            ac $logfile "Detected user: $($userURL)`n"  
            $user_detected = $True  
        }  
    }until($user_detected)  
}else{  
    ac $logfile "Mapping to a sharepoint library, not a user library"  
    $browseURL = $sharepointURL  
}  
  
$ie.navigate($browseURL)  
do {sleep -m 100} until (-not ($ie.Busy)) 
sleep 1  
do {sleep -m 100} until (-not ($ie.Busy)) 
  
#don't do the URL check if going to Sharepoint, since it'll add a lot of crap we won't like  
if($ie.LocationURL.StartsWith($baseURL + "/personal/" + $userURL) -ne $True -and $sharepointURL.Length -lt 1){  
    ac $logfile "failed to detect expected URL: $($browseURL), Actual URL: $($ie.LocationURL)" 
}else{  
    ac $logfile "Current location: $($ie.LocationURL)"  
    ac $logfile "Session established, attempting to map drive`n"  
    $mapresult = MapDrive $driveLetter $mapURL $driveLabel  
}  
ac $logfile "Script done, exiting`n"  
abort_OM
 
