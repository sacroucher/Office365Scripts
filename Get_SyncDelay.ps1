<# 
    .SYNOPSIS 
        Get the amount of time for an update to Office365 
    .DESCRIPTION 
        This script makes a change to an account and then checks to see how long 
        it takes before that change is replicated on Office365. 
    .PARAMETER Tenant 
        This is your tenant name 
    .PARAMETER Account 
        This is the test account name 
    .PARAMETER Credential 
        This is a tenant account with rights to modify user objects 
    .EXAMPLE 
        .\Get-SyncDelay.ps1 -UserName tenantadmin@myTenant.onmicrosoft.com -Password mYP@ssw0rd -Tenant myTenant.onmicrosoft.com -Account tstAcct01 
 
        Get-LogFile -LogName Office365-Sync-Check |Where-Object -Property EventID -eq 0 
 
        EventRecord : 00000003 
        LogName     : Office365-Sync-Check 
        Source      : Set-MsolUser 
        Time        : 03/03/2015 17:39:42 
        EventID     : 0 
        EntryType   : Information 
        Message     : Total Sync Delay in Seconds 0 
     
        Description 
        ----------- 
        This example shows using the script by passing in a valid admin account and password 
        as strings, useful for certain types of automation. The Get-LogFile cmdlet from LogFiles 
        was used to get the delay. 
    .EXAMPLE 
        $Credential = New-Object System.Management.Automation.PSCredential ('tenantadmin@myTenant.onmicrosoft.com', (ConvertTo-SecureString -String 'mYP@ssw0rd' -AsPlainText -Force)); 
        .\Get-SyncDelay.ps1 -Credential $Credential -Tenant myTenant.onmicrosoft.com -Account tstAcct01 
 
        Starting Sync Check 
        Change DisplayName property of tstAcct01@myTenant.onmicrosoft.com to Delay Monitor 1425404038 
 
 
        Days              : 0 
        Hours             : 0 
        Minutes           : 0 
        Seconds           : 1 
        Milliseconds      : 78 
        Ticks             : 10781009 
        TotalDays         : 1.24780196759259E-05 
        TotalHours        : 0.000299472472222222 
        TotalMinutes      : 0.0179683483333333 
        TotalSeconds      : 1.0781009 
        TotalMilliseconds : 1078.1009 
     
        Description 
        ----------- 
        This example shows passing in a credential object which is useful for running the 
        script manually. 
    .NOTES 
        This came from code that was provided to me originally through the 
        Office365 HiEd mailling list.  
 
        Many thanks to Jacob Fortune for sharing! 
 
        In order for this script to work you will need the MSONline Module 
        that you can download from the first link. You will also need my 
        LogFiles module that you can download from the second link. If you would 
        rather write to the Windows EventLog, you can easily make that modification 
        yourself, just remember to remove the Import-Module Logfiles statement 
        in the Begin section. 
    .LINK 
        https://msdn.microsoft.com/en-us/library/azure/jj151815.aspx 
    .LINK 
        https://raw.githubusercontent.com/jeffpatton1971/mod-posh/master/powershell/production/includes/LogFiles.psm1 
#> 
Param 
( 
    [Parameter(Mandatory=$true,ParameterSetName='Credential')] 
    [System.Management.Automation.PSCredential]$Credential, 
    [Parameter(Mandatory=$true,ParameterSetName='NoCredential')] 
    [string]$UserName, 
    [Parameter(Mandatory=$true,ParameterSetName='NoCredential')] 
    [string]$Password, 
    [Parameter(Mandatory=$true)] 
    [string]$Tenant, 
    [Parameter(Mandatory=$true)] 
    [string]$Account, 
    [Parameter(Mandatory=$false)] 
    [switch]$ExchangeOnline 
) 
Begin 
{ 
    # 
    # Import the MSONLINE module 
    # Import the LogFiles module 
    # 
    $LogName = "Office365-Sync-Check"; 
    try 
    { 
        Import-Module MSOnline; 
         
        if ($PSCmdlet.ParameterSetName -eq 'NoCredential') 
        { 
            $Credential = New-Object System.Management.Automation.PSCredential ($UserName, (ConvertTo-SecureString -String $Password -AsPlainText -Force)); 
            Import-Module C:\projects\mod-posh\PowerShell\Production\Includes\LogFiles.psm1; 
            } 
 
        if ($ExchangeOnline) 
        { 
            # 
            # EOP Standalone URL  
            # $EoURI = "https://ps.protection.outlook.com/powershell-liveid" 
            # 
            # Exchange Online Tenant 
            $EoUri = "https://outlook.office365.com/powershell-liveid" 
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $EoUri -Credential $Credential -Authentication Basic -AllowRedirection; 
            # 
            # Import ExchangeOnlline Cmdlets 
            # 
            Import-PSSession $Session; 
            }         
        Connect-MsolService -Credential $Credential; 
        } 
    catch 
    {         
        $LogMessage = $Error[0].Exception; 
        if ($PSCmdlet.ParameterSetName -eq 'Credential') 
        { 
            Write-Error $LogMessage; 
            } 
        else 
        { 
            Write-LogFile -LogName $LogName -Source "Script Init" -EventId 1 -EntryType "Error" -Message $LogMessage; 
            } 
        break; 
        } 
    } 
Process 
{ 
    try 
    { 
        if ($PSCmdlet.ParameterSetName -eq 'Credential') 
        { 
            Write-Host "Starting Sync Check"; 
            } 
        else 
        { 
            Write-LogFile -LogName $LogName -Source "Begin Sync Check" -EventId 10 -EntryType "Information" -Message "Starting Sync Check"; 
            } 
        $TimeStamp = "{0:G}" -f [int][double]::Parse((Get-Date -UFormat %s)); 
        $UserPrincipalName = "$($Account)@$($Tenant)"; 
        $DisplayName = "Delay Monitor $($TimeStamp)"; 
         
        $StartTime = Get-Date; 
        Set-MsolUser -UserPrincipalName $UserPrincipalName -DisplayName $DisplayName; 
         
        if ($PSCmdlet.ParameterSetName -eq 'Credential') 
        { 
            Write-Host "Change DisplayName property of $($UserPrincipalName) to $($DisplayName)"; 
            } 
        else 
        { 
            Write-LogFile -LogName $LogName -Source "Set-MsolUser" -EventId 10 -EntryType "Information" -Message "Change DisplayName property of $($UserPrincipalName) to $($DisplayName)"; 
            } 
 
        if ($ExchangeOnline) 
        { 
            $MsolUser = Get-Mailbox -Identity $UserPrincipalName |Select-Object -Property DisplayName, WhenChanged;  
            } 
        else 
        { 
            $MsolUser = Get-MsolUser -UserPrincipalName $UserPrincipalName |Select-Object -Property DisplayName, WhenChanged; 
            } 
        while ($MsolUser.DisplayName -ne $DisplayName) 
        {     
            if ($PSCmdlet.ParameterSetName -eq 'Credential') 
            { 
                Write-Host "Getting DisplayName property of $($UserPrincipalName)"; 
                } 
            else 
            { 
                Write-LogFile -LogName $LogName -Source "Get-MsolUser" -EventId 10 -EntryType "Information" -Message "Getting DisplayName property of $($UserPrincipalName)"; 
                } 
            Start-Sleep -Seconds 5; 
            if ($ExchangeOnline) 
            { 
                $MsolUser = Get-Mailbox -Identity $UserPrincipalName |Select-Object -Property DisplayName, WhenChanged;  
                } 
            else 
            { 
                $MsolUser = Get-MsolUser -UserPrincipalName $UserPrincipalName |Select-Object -Property DisplayName, WhenChanged; 
                } 
            } 
         
        $EndTime = Get-Date; 
        $Delay = New-TimeSpan -Start $StartTime -End $EndTime 
 
        if ($PSCmdlet.ParameterSetName -eq 'Credential') 
        { 
            return $Delay; 
            } 
        else 
        { 
            Write-LogFile -LogName $LogName -Source "Set-MsolUser" -EventId 0 -EntryType "Information" -Message "Total Sync Delay in Seconds $($Delay.Seconds)"; 
            } 
        } 
    catch 
    { 
        $LogMessage = $Error[0].Exception; 
        if ($PSCmdlet.ParameterSetName -eq 'Credential') 
        { 
            Write-Error $LogMessage; 
            } 
        else 
        { 
            Write-LogFile -LogName $LogName -Source "Script Process" -EventId 1 -EntryType "Error" -Message $LogMessage; 
            } 
        break; 
        } 
    } 
End 
{ 
    } 