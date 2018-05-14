#Import the Modules we needs#
Import-Module activedirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue

$GroupMembers = Get-ADGroupMember -Identity 'TEST GROUP'
$SaveDate = [string](Get-Date -Format yyyyMMdd_HHmm)
$Path = 'C:\EMS\Project_X_' + $SaveDate + '.csv'
$CustomList = @()


ForEach ($GroupMember in $GroupMembers){

    $WSName = Get-ADUser -Identity $GroupMember -Pr * | Select-Object -ExpandProperty extensionAttribute10 -First 1
    Write-Output $WSName
#    Method1    Stop Printer Services - Working
    Invoke-Expression "psservice64.exe -nobanner \\$WSName stop spooler"
    Invoke-Expression "psservice64.exe -nobanner \\$WSName setconfig spooler disabled"


    #Set the user's Mailbox to 0 Send/Receive
    $SamName = Get-ADUser -Identity $GroupMember -Pr * | Select-Object -ExpandProperty CN -First 1
    Write-Output $SamName

    #Get the Information about the user
    $MagicMailbox = ForEach ($User in $SamName) {
    $MBX = Get-Mailbox -Identity $User
    $MBX = $MBX | Select-Object -Property Name,UseDatabaseQuotaDefaults,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,ServerName,Database
    $CBX = Get-CASMailbox -Identity $User
    $CBX = $CBX | Select-Object -Property Name,OWAEnabled, PopEnabled,ImapEnabled,ActiveSyncEnabled

    $CustomObject = New-Object -TypeName PSObject -Property (@{
        'UsersName' = $MBX.Name;
        'DBdefault' = $MBX.UseDatabaseQuotaDefaults;
        'IssueWarn' = $MBX.IssueWarningQuota;
        'ProhibitS' = $MBX.ProhibitSendQuota;
        'ProhibitR' = $MBX.ProhibitSendReceiveQuota;
        'ServerNam' = $MBX.ServerName;
        'DatabaseX' = $MBX.Database;
        'OWAEnable' = $CBX.OWAEnabled
        'PopEnable' = $CBX.PopEnabled;
        'ImpEnable' = $CBX.IMAPEnabled;
        'ActEnable' = $CBX.ActiveSyncEnabled;
        })
    $CustomList += $CustomObject
}
$CustomList | Select-Object -Pr * | Export-Csv -Path $Path -NoTypeInformation -NoClobber -Force

    #Set the User's Mailbox to Zero
    #Set-Mailbox -Identity $SamName -IssueWarningQuota 0 -ProhibitSendQuota 0 -ProhibitSendReceiveQuota 0

    #Disconnect User's Mailbox
    #Disable-Mailbox -Identity $SamName

# Set Proxy to stop all Internet Usage, but leave networking functional.
#reg delete\\$WSName\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable
#reg delete\\$WSName\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride
#reg delete\\$WSName\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer

}

$TestList = $CustomList | Select -First 1

Foreach ($Obj in $TestList){

    Set-Mailbox -Identity $Obj.UsersName UseDatabaseQuotaDefaults
    Set-Mailbox -Identity $Obj.UsersName -IssueWarningQuota 0 -ProhibitSendQuota 0
    Get-MailboxStatistics -Server $Obj.ServerNam | where { $_.DisconnectDate -ne $null } | select DisplayName,DisconnectDate
    Connect-Mailbox -database $Obj.DatabaseX -Identity $Obj.UsersName
    Set-Mailbox -Identity $Obj.UsersName -UseDatabaseQuotaDefaults $false
    
    }
