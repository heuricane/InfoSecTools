#
# Might be easier to contol services than one thinks
#
$target="ZOLTAN"
$svc = "Spooler"
Get-Service -ComputerName $target -Name $svc | Stop-Service
#
Get-Service -ComputerName $target -Name $svc | Start-Service
#
#
# What happens when you try this?
#
$user = "Jay.Berko.CTR@snail.mil"
$dc = $env:LOGONSERVER
Disable-Mailbox -Identity $user -DomainController $dc -Confirm
#
Connect-Mailbox -Identity $user -DomainController $dc -Confirm
#
#
#Temporarily disable internet access tough to turn back on so we'll invoke the command with a timer
#If you don't like the timer, talk to networking for enabling/disabling port.
#
$svc = "Netman"
$timer = 600 #10 minutes
$sesh = New-PSSession -ComputerName $target
Invoke-Command -Session $sesh {
    Get-Service -ComputerName $target -Name $svc | Stop-Service
    Start-Sleep -Seconds $timer
    Get-Service -ComputerName $target -Name $svc | Start-Service
    }
#
#
# Combining data
# First of all, long commands suck, here's a trick to shorten
#
$WSName = Get-ADUser -Identity $GroupMember -Properties * | Select-Object extensionAttribute10
$WSName = $WSName | Select-Object -ExpandProperty extensionAttribute1 -First 1
Write-Output $WSName



# Now, I don't see where you set $SamName, I assume it's not one at a time.
# This example assumes $SamName is a list we loop through

$date = Get-Date -format d
$savedate = (Get-Date).tostring("yyyyMMdd")
$path = 'C:\EMS\' + $savedate + '.csv'

$CustomList = @()

Foreach ($SamName in $SamList){

    $mbx = Get-Mailbox -Identity $SamName
    $mbx = $mbx | Select -Pr Name,UseDatabaseQuotaDefaults,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,ServerName,Database 
    $mbx = $mbx | Sort-Object Name,ServerName

    $cbx = Get-CASMailbox -Identity $SamName 
    $cbx = $cbx | Select-Object -Property OWAEnabled, PopEnabled,ImapEnabled,ActiveSyncEnabled 
    $cbx = $cbx | Sort-Object Name,ServerName

    $CustomObject = New-Object -TypeName PSObject -Property (@{
        'UsersName' = $mbx.Name
        'DBdefault' = $mbx.UseDatabaseQuotaDefaults;
        'IssueWarn' = $mbx.IssueWarningQuota;
        'ProhibitS' = $mbx.ProhibitSendQuota;
        'ProhibitR' = $mbx.ProhibitSendReceiveQuota
        'ServerNam' = $mbx.ServerName
        'Databasex' = $mbx.Database
        'OWAEnable' = $cbx.OWAEnabled
        'PopEnable' = $cbx.PopEnabled
        'ImpEnable' = $cbx.IMAPEnabled
        'ActEnable' = $cbx.ActiveSyncEnabled
        })

    $CustomList += $RomysCustomObject

}

$CustomList | Export-Csv -Path $path -NoTypeInformation
