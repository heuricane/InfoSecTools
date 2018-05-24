$currentversion = "Version 1.96"
#
#  WATCH SCRIPT 1.96
#  2013 AUG 11
#  JAY C BERKOVITZ
#     
#  To alter servers in sites, scroll down below the README section to the Environmentals for variables
#
#  README / Help menu function
#
Function helpmenu {
    $HelpText = @"
    
    There are seven functions to date, 
    and after each one an option to return to main menu
    Don't worry, this script is harmless and there's always a way out.
    by typing 'no' or the infamous control+C, or closing the window

    *MS Exchange Checks*
    Function is simple enough:
    Sends 'Get-Queue' to EH11 and EH12 for message count and status
    then a 'Get-StorageGroupCopyStatus' to UMEBAEM12 
    an 'if' statement with a 'send-message' tests the BB server
    Then there are arrays for server types that roll through a 
    'get-service' command to call for displayname and status

    *Defragment and Analyzer*
    First, it asks for a site or machine name and you choose
    on a side note, the single target machine option doesn't work yet
    so you chose jeb, isa, or nsa and it sends that array onward
    'C:\'.DefragAnalysis().DefragRecommended offers a suggestion 
    for each server in the selected site (be it true or false)
    the user can choose to defrag all, none, or only recommended
    Drive.ChkDsk() does a default defrag of all selected targets in site
    
    *Pre-Reboot Cleanup*
    You choose a site or machine name (but single target option works here)
    The server arrays are the same as before, isa / jeb / nsa 
    it lists the targets in array and asks for confirmation
    Then it clears all unlocked items from \Windows\temp and all Hotfix logs
    Then appears to hang when scanning/deleting profiles over 180 days old
    
    *Reboot Monitor*
    Ok, I can't take any credit for this one, googling around 
    found here - http://gallery.technet.microsoft.com/scriptcenter
    /2d537e5c-b5d4-42ca-a23e-2cbce636f58d  now it's not perfect.
    but it will work for now.  the timer is 5 seconds and unfortunately
    there is no way back out of it.  You have to close the window
   
    *Reboot Initiator*
    Simplest function of them all.  Enter server name, hit enter. again?
    
    *Post-Reboot Service Checks*
    this time all three site arrays are combined into one array allsite
    allsite | foreach {Get-WmiObject -Class Win32_Service -Computer _
    -Filter StartMode=Auto and state=Stopped name!=exceptionslist}
    It checks all servers services for automatically started
    services that are currently in the stopped state using get-wmiobject
    
    *VMware Power CLI Window*
    Not really a function, just opens a VMwareCLI window with advice
    
    *Commvault Storage Utilization Report*
    Imports most recent LibraryandDriveReport from UMEBABU11, copies it to D:
    Displays partial data, then asks if you want to email to ISWO watch.
    
"@
    $HelpText
}
#
#    Set Environmentals
#
$mx = "smtp-int.me.navy.mil"
$site3 = @(
'Server21'
'Server22'
'Server23')
$site2 = @(
'Server11'
'Server12'
'Server13')
$site1 = @(
'Server11'
'Server12'
'Server13')
$site4 = @(
'Server31')
$MEMailBoxServers = @(
'JAYSERVEREM11'
'JAYSERVEREM12')
$allsite = $site1 + $site2 + $site3 + $site4
$sitechoice = @"    
    |    ---------------------------
    |    Type '1' for Site 1
    |    Type '2' for Site 2
    |    Type '3' for Site 3
    |    Type '4' for Site 4
    |    Type 'a' for ALL the Sites
    | Or Type hostname of target
    |
    |
"@
#
#  List of automatically services that are not critical to functionality
#
$svcfilter="StartMode='Auto' and state='Stopped' and name!='VSS' and name!='swprv' and name!='SwdisRestart' and name!='SysmonLog' and name!='sppsvc' and name!='ShellHWDetection' and name!='WinRM' and name!='MSIServer' and name!='vmvss' and name!='W32Time' and name!='TBS' and name!='clr_optimization_v4.0.30319_32' and name!='clr_optimization_v4.0.30319_64' and name!='SMSMSE' and name!='RoxLiveShare10' and name!='wscsvc'"


$runonreplay="n"
#
#  Method to derive email address from admin account name
#
  $whoami = "$env:username"
  If ($whoami -like '*y.berkov*'){$thisguy = "jay.berkovitz.ctr"}
  If ($whoami -like 'Jay.berko*'){$thisguy = "jay.berkovitz.ctr"}
  $thisguy = $thisguy + "@myaddress.com"
#
#  Technet function to get Share access stats
#



function Set-Color{
param([String] $Color = $(throw "Please specify a color."))
# Trap the error and exit the script if the user
# specified an invalid parameter.
trap [System.Management.Automation.RuntimeException] {
  write-error -errorrecord $ERROR[0]
  exit
}
# Assume -color specifies a hex value and it cast to a [Byte].
$newcolor = [Byte] ("0x{0}" -f $Color)

# Split the color into background and foreground colors.
# [Math]::Truncate returns a [Double], so cast it to an [Int].
$bg = [Int] [Math]::Truncate($newcolor / 0x10)
$fg = $newcolor -band 0xF

# If the background and foreground match, throw an error;
# otherwise, set the colors.
if ($bg -eq $fg) {
  write-error "The background and foreground colors must not match."
} else {
  $HOST.UI.RawUI.BackgroundColor = $bg
  $HOST.UI.RawUI.ForegroundColor = $fg
}
}
#
#    Share Connection
# 
function Get-ShareConnection{ 
  Param 
  ( 
      # param1 help description 
      [Parameter(Mandatory=$true,  
                 ValueFromPipeline=$true, 
                 ValueFromPipelineByPropertyName=$true,  
                 ValueFromRemainingArguments=$false,  
                 Position=0)] 
      [Alias("cn")]  
      [String[]]$ComputerName 
  ) 
  Begin 
  { 
  } 
  Process 
  { 
    $ComputerName | ForEach-Object { 
    $Computer = $_ 
    try { 
      Get-WmiObject -class Win32_ConnectionShare  -namespace root\cimv2 -comp $Computer -EA Stop |  
      Group-Object Antecedent | 
      Select-Object @{Name="ComputerName";Expression={$Computer}}, 
                    @{Name="Share"       ;Expression={(($_.Name -split "=") |  
                    Select-Object -Index 1).trim('"')}}, 
                    @{Name="Connections" ;Expression={$_.Count}}
      } 
      catch  
        { 
          "    Cannot connect to $Computer"  
        } 
     }#ForEach-Object 
  } 
  End 
  { 
  } 
}

function exitscript {
   Set-Color 0F
   Clear-Host
   exit
}
#
#    Wait Job
#
function WaitJob($job){
	#$saveY = [console]::CursorTop
	#$saveX = [console]::CursorLeft   
	$str = ' '
	$date = ( get-date -format "dddd, MMMM dd, yyyy HH:MM:ss").ToCharArray()
	$saveYinit = [console]::CursorTop
	$saveXinit = [console]::CursorLeft      
	$Start = (get-date).AddHours(2)
	do {
		$saveX = $saveXinit
		$saveY = $saveYinit
		[console]::setcursorposition($saveX,$saveY)
		$date = get-date
		$output = "Next run: $Start  Time Now: $date"
        $secRemaining = ($start - $date).TotalSeconds
		Write-Progress -Activity "Waiting for next checks" -SecondsRemaining $secRemaining -status $output 
		$State = (Get-Job -Name $job).state -eq 'Running'
		if ($state) {$running = $true}
		else {$running = $false}
	} while ($running)
	Remove-Job $job
    Write-Progress -Activity "Waiting for next checks" -SecondsRemaining $secRemaining -status $output -Completed 
} # end function

#
#    Comm Vault Report
#
function CommVaultReport{

	$mx = "smtp-int.myaddress.com"
	$Server = "MyBackupServer"
	$Path = "\\$server\C`$\Reports"
	$ReportName = "GalaxyReport_libraryandDriveReport*.csv"
	$ReportFile = get-childitem $Path -Filter $ReportName | Where-Object {($_.CreationTime).date -eq (get-date).date}
	$Report = Get-content $ReportFile.FullName
	$md11Info = $Report | Where-Object {$_ -match "MY_Dell_MD_Server11"} 
	$md11Info = $md11Info[0].split(',')
	$md11Capacity = $md11Info[1]
	$md11SpaceLeft = $md11Info[2]
	$md11UsedSpace = $md11Info[3]
	$md11RemainingPerc = "{0:P2}" -f ($md11SpaceLeft / $md11Capacity)
	$md11UsedPerc = "{0:P2}" -f ($md11UsedSpace / $md11Capacity)

	$md12Info = $Report | Where-Object {$_ -match "MY_Dell_MD_Server12"} 
	$md12Info = $md12Info[0].split(',')
	$md12Capacity = $md12Info[1]
	$md12SpaceLeft = $md12Info[2]
	$md12UsedSpace = $md12Info[3]
	$md12RemainingPerc = "{0:P2}" -f ($md12SpaceLeft / $md12Capacity)
	$md12UsedPerc = "{0:P2}" -f ($md12UsedSpace / $md12Capacity)
	$ReportFileName = $ReportFile.Name
	$Body =@"
	For Official Use Only
	`n
	Backup Server: $Server
	Repot File: $ReportFileName
	File Location: $Path
	`n
	NIPR`tCapacity`tSpace Left`t% Remaining`t`t% Used
	MD11`t$md11Capacity`t$md11SpaceLeft`t$md11RemainingPerc`t`t$md11UsedPerc
	MD12`t$md12Capacity`t$md12SpaceLeft`t$md12RemainingPerc`t`t$md12UsedPerc
"@

	$date = get-date -format "dd MMM yyyy"
	$Subject = "Daily Utilization Report $date"
	$email = @{
				From = $thisguy
				To = "JayBerkovitz@myaddress.com"
				Subject = $Subject
				SMTPServer = $mx
				Body = $body
			}
		Send-MailMessage @email
}
#
#    Analyze/Defrag Function
#
function defragglerock {
   echo @"
    |
    |
    |    Analyze/Defragment Function
"@
   $sitechoice
   $defragwho = read-host "    |   Enter name of site or server"
   If ($defragwho -eq "4"){$fragged = $siteavu}
    elseIf ($defragwho -eq "3"){$fragged = $sitejeb}
     elseIf ($defragwho -eq "2"){$fragged = $siteisa} 
       elseIf ($defragwho -eq "1"){$fragged = $sitensa}
         elseIf ($defragwho -eq "a"){$fragged = $allsite}
           else {$fragged = $defragwho}
   echo @"
    |
    |   List of Servers in array:
    |
"@
   $fragged | foreach {
      echo  "    |    - $_ "
      }
   echo "    |"
   $confirmanal = read-host "    |   Continue with analysis? [y/n]"
   echo "    |"   
      If ($confirmanal -eq "y")
      {   
          $fragged | foreach {
              $Drives = Get-WMIObject Win32_Volume –filter "DriveLetter='C:'" –computer $_
              foreach($drive in $drives)
              {
                  $ShouldDefrag = $Drive.DefragAnalysis().DefragRecommended
                  write-Host -fore yellow "    |  $_ Defragmentation Recommended?: $ShouldDefrag"
              }
          }
      echo "    |"
      $choosewho = read-host "    |  Do you want to defrag recommended [y/n]"
            if ($choosewho -eq "y")
            {
               if($ShouldDefrag)
               {  
                  echo "    |"
                  echo "    |  ...Defragmentation in progress..."
                  $DefragResult = $drive.Defrag($true)
                  if($DefragResult -eq "") {$Drive.ChkDsk($false,$true,$true,$false,$false,$false)}
                  echo "    |"
                  echo "    |  Defragementation of C: on $servers complete"
               }
            }
   }
   echo "    |"
}
#
#    PreReboot Cleanup Function
#
function callthecleaner {
    echo @"
    |
    |
    |    Pre-Reboot Cleanup Function
"@
    $sitechoice
    $dirty = read-host "    |   Enter name of site or server"
    If ($dirty -eq "4"){$cleanlist = $siteavu}
        elseIf ($dirty -eq "3"){$cleanlist = $sitejeb}
            elseIf ($dirty -eq "2"){$cleanlist = $siteisa} 
                elseIf ($dirty -eq "1"){$cleanlist = $sitensa}
                    elseIf ($dirty -eq "a"){$cleanlist = $allsite}
                        else {$cleanlist = $dirty}
    echo @"
    |
    |   List of Servers in array:
    |
"@
    $cleanlist | foreach {echo "    |    - $_ "}
    echo @"
    |
    |  This script will delete the following:
    |  
    |    - All items in C:\Windows\Temp\
    |    - All items in C:\Windows that begin with the letters 'KB2'
    |    - All folders in C:\Windows that begin with a '$'
    |    - All profiles older than 180 days of age
    |
    |
"@
    $confirmclean = read-host "    |  Continue with cleanup? [y/n]"
    echo "    |"   
    if ($confirmclean -eq "y")
    {
        $cleanlist | foreach {
            $thisos = Get-WmiObject -Class Win32_OperatingSystem -Namespace root/cimv2 -ComputerName $_ | select-object Name
            echo "Operating System: $thisos"
            echo "    |"
            remove-item \\$_\c$\windows\temp\* -Recurse -ErrorAction SilentlyContinue
            remove-item \\$_\c$\windows\KB* -ErrorAction SilentlyContinue
            remove-item \\$_\c$\windows\$* -Recurse -ErrorAction SilentlyContinue
            echo "    |    $_ Has been Cleaned"
            echo "    |"
        }
        If ($thisos -like "*2003*")
        {
            echo "    |   ...Scanning/Deleting profiles over 180 days stale... "
            echo "    |"
            foreach ($dirtyserver in $cleanlist) {\\MyTerminlaServer\C$\AdminScripts\Tools\DelProf2.exe /q /c:$dirtyserver /d:180}
        }
#  
#  This part is Kayode's, but only in part, 
#  because I don't wanna have to run script from 2008 server for the Win32_UserProfile function
#        
        Elseif ($thisos -like "*2008*")
        {
            foreach ($dirtyprofiles in $cleanlist) 
            {
            $UserList = gwmi -class Win32_UserProfile -comp $dirtyprofiles | 
            where-object {$_.ConvertToDateTime($_.lastusetime) -lt (get-date).addDays(-"365") -and $_.Special -eq $False}
            $Output = $UserList | select @{label="last used";EXPRESSION={$_.ConvertToDateTime($_.lastusetime)}},LocalPath, SID | ft -a
            if ($UserList -eq $Null)
               {echo "    |   No profiles found!"y
        	    return}
            else {$Output}
            $in = Read-Host "    |   Are you sure you want to delete these profiles? (Yes/No)"
            if ($in -eq "Yes") {$UserList | foreach-object {$_.delete()}}
            else {echo "    |   No Profiles deleted."}
            }
        }
        echo "    |   Cleanup sequence completed"
    }
}
#
#     Reboot Monitorer, stolen off Microsofts technet
#
function start-monitor {      
[CmdletBinding()]
 Param              
    (                        
    [Parameter(Mandatory=$false, 
               Position=0,                          
               ValueFromPipeline=$true,             
               ValueFromPipelineByPropertyName=$true)] 
    [String[]]$ComputerName,         
    # reset the lists of hosts prior to looping 
    $OutageHosts = @(), 
    # specify the time you want email notifications resent for hosts that are down 
    $EmailTimeOut = 30, 
    # specify the time you want to cycle through your host lists. 
    $SleepTimeOut = 2, 
    # specify the maximum hosts that can be down before the script is aborted 
    $MaxOutageCount = 10, 
    # specify who gets notified 
    $notificationto = "JayBerkovitz@myaddress.com", 
    # specify where the notifications come from 
    $notificationfrom = "JayBerkovitz@myaddress.com" 
    )#End Param 
echo @"
    |
    |
    |      Connectivity Monitor
"@
$sitechoice
$monitorwho = read-host "    |   Enter name of site or server"
   If ($monitorwho -eq "4"){$ComputerName = $site4}
    elseIf ($monitorwho -eq "3"){$ComputerName = $site3}
     elseIf ($monitorwho -eq "2"){$ComputerName = $site2} 
       elseIf ($monitorwho -eq "1"){$ComputerName = $sit1}
         elseIf ($monitorwho -eq "a"){$ComputerName = $allsite}
           else {$ComputerName = $monitorwho}
# start looping here
do
{ 
   $available = @() 
   $notavailable = @() 
   Write-Host (Get-Date) 
   # Read the File with the Hosts every cycle, this way to can add/remove hosts 
   # from the list without touching the script/scheduled task,  
   # also hash/comment (#) out any hosts that are going for maintenance or are down. 
   $ComputerName | Where-Object {!($_ -match "#")} |  
   #"test1","test2" | Where-Object {!($_ -match "#")} | 
   ForEach-Object { 
      if(Test-Connection -comp $_ -Count 1 -ea silentlycontinue) 
      {
         # if the Host is available then write it to the screen 
         write-host "    |    Available host ---> "$_ -BackgroundColor Green -fore black 
         [Array]$available += $_ 
         # if the Host was out and is now backonline, remove it from the OutageHosts list 
         if ($OutageHosts -ne $Null) 
         {
            if ($OutageHosts.ContainsKey($_)) 
            {
               $OutageHosts.Remove($_)          
            }
         }
      }
      else
      {
         # If the host is unavailable, give a warning to screen 
         write-host "    |   Unavailable host ------------> "$_ -BackgroundColor Magenta -ForegroundColor White 
         if(!(Test-Connection -comp $_ -Count 2 -ea silentlycontinue)) 
         { 
            # If the host is still unavailable for 4 full pings, write error and send email 
            write-host "        |   Unavailable host ------------> "$_ -BackgroundColor Magenta -ForegroundColor White 
            [Array]$notavailable += $_ 
            if ($OutageHosts -ne $Null) 
            { 
                if (!$OutageHosts.ContainsKey($_)) 
                { 
                   # First time down add to the list and send email 
                   Write-Host "$_ Is not in the OutageHosts list, first time down" 
                   $OutageHosts.Add($_,(get-date)) 
                   $Now = Get-date 
                   #$Body = "$_ has not responded for 5 pings at $Now" 
                   #Send-MailMessage -Body "$body" -to $notificationto -from $notificationfrom ` 
                   # -Subject "Host $_ is down" -SmtpServer $smtpserver 
                } 
                else 
                { 
                   # If the host is in the list do nothing for 1 hour and then remove from the list. 
                   Write-Host "$_ Is in the OutageHosts list" 
                   if (((Get-Date) - $OutageHosts.Item($_)).TotalMinutes -gt $EmailTimeOut) 
                   {$OutageHosts.Remove($_)} 
                } 
            } 
         else 
            { 
               # First time down create the list and send email 
               Write-Host "Adding $_ to OutageHosts." 
               $OutageHosts = @{$_=(get-date)} 
               #$Body = "$_ has not responded for 5 pings at $Now"  
               #Send-MailMessage -Body "$body" -to $notificationto -from $notificationfrom ` 
               # -Subject "Host $_ is down" -SmtpServer $smtpserver 
            }  
         } 
      } 
   } 
   # Report to screen the details 
   write-host "Available count:"$available.count 
   write-host "Not available count:"$notavailable.count 
   write-host "Not available hosts:" 
   $OutageHosts 
   write-host "" 
   write-host "Sleeping $SleepTimeOut seconds" 
   sleep $SleepTimeOut 
   if ($OutageHosts.Count -gt $MaxOutageCount) 
   { 
      # If there are more than a certain number of host down in an hour abort the script. 
      $Exit = $True 
      $body = $OutageHosts | Out-String 
      Send-MailMessage -Body "$body" -to $notificationto -from $notificationfrom -Subject "More than $MaxOutageCount Hosts down, monitoring aborted" -SmtpServer $mx 
   } 
} 
while ($Exit -ne $True) 
}
#
#     Reboot Servers Function
#
function rebootserver {
    $zombiecount=0
    $zomebiekill=0
    $theundead=@()
    echo @"
    |
    |
    |    Initiate Reboot Function

"@
    $sitechoice
    $zombielist = read-host "    |   Enter name of site or server"
    If ($zombielist -eq "4"){$rebootwho = $site4}
        elseif ($zombielist -eq "3"){$rebootwho = $site3}
            elseIf ($zombielist -eq "2"){$rebootwho = $site2} 
                elseIf ($zombielist -eq "1"){$rebootwho = $site1}
                    elseIf ($zombielist -eq "a"){$rebootwho = $allsite}
                        else {$rebootwho = $zombielist}
   echo @"
    |
    |   List of Servers in array:
    |
"@
    $rebootwho | foreach {echo "    |    - $_ "}
    foreach ($zombie in $rebootwho){$zombiecount++}
    echo " "
    $necromancy = read-host "    |  Are you certain? [y/n]"
    if ($necromancy = "y")
    {
        $rebootwho | foreach {
            if (test-connection $_ -count 1 -ea silentlycontinue)
            {
                write-host -fore green " - Send reboot to $_"
                echo " "
                Restart-computer -comp $_ -throttlelimit -1
                $theundead += $_
                $risen=0
            }
            do {
                if (test-connection $_ -count 1 -ea silentlycontinue)
                {
                    echo "    Waiting for $_ to die..."
                    $rising=1
                }
                else 
                {
                    do {
                        if (test-connection $_ -count 1 -ea silentlycontinue)
                        {
                            $rising=0
                            $risen=1
                            $zombiekill ++
                            echo " "
                            write-host -fore yellow "    Happy Easter"                                 
                            echo " "
                        }
                        else 
                        {
                            $rising=1
                            echo "    $_ died, awaiting the resurrection..."
                            sleep 5
                        }
                    }
                    while ($rising -eq 1)
                }
                sleep 5
            } 
            while ($risen -eq 0)
        }
        echo " "
        echo "Reboots sent to $theundead"
        echo " "
        write-host -fore yellow "    $zombiekill zombies resurrected"
        echo " "
    }    
}
#
#    Check for Stopped Services Function
#
function autostopped {
    echo @"
    |
    |
    |     Post Reboot Service Check
"@
    $sitechoice
    write-host -fore yellow "    |   - Only displays failed services -"
    echo "    |"
    echo "    |"
    $siteservices = read-host "    |   Enter name of site"
    echo " "
    If ($siteservices -eq "4"){$targetsvcs = $site4}
        elseIf($siteservices -eq "3"){$targetsvcs = $site3}
            elseIf ($siteservices -eq "2"){$targetsvcs = $site2} 
                elseIf ($siteservices -eq "1"){$targetsvcs = $site1}
                    elseIf ($siteservices -eq "a"){$targetsvcs = $allsite}
    $targetsvcs | foreach {
        echo "    |   $_ "
        Get-WmiObject -Class Win32_Service -Comp $_ -Filter $svcfilter |
        select displayname,state,startmode | Out-String
    }
    echo "    |"
}
#
#    Blackberry Check 
# 
Function BlackberryCheck{

Send-MailMessage -to "Admin1@myaddress.com" -from 'Admin2@MyAddress.com' -subject '<$Confirm>' -smtpserver "smtp-int.MyAddress.com"
"Blackberry email confirmation message sent`n" 
}

#
#    Exchange Check
# 
function ExchangeCheck {
    #
    # Microsoft Exchange Management Powershell Snapin
    #
    $s = "Microsoft.Exchange.Management.PowerShell.Admin"
    if (get-pssnapin $s -ea "silentlycontinue") {
    }
    elseif (get-pssnapin $s -registered -ea "silentlycontinue") {
        #"PSsnapin $s is registered but not loaded"
        Add-PSSnapin $s
    }
    else {
         "PSSnapin $s not found. Please install Exchange 2007 Management Tools and try again." 
        Break
    }
	"****EH 11 & EH 12 Queues****" 
	Get-Queue -Server Exchange_HUB_Server | select DeliveryType,Status,MessageCount,NextHopDomain | Format-Table -AutoSize  | Out-String
	Get-Queue -Server Exchange_HUB_Server | select DeliveryType,Status,MessageCount,NextHopDomain | Format-Table -AutoSize | Out-String
    Get-Queue -Serve  Exchange_HUB_Server | select DeliveryType,Status,MessageCount,NextHopDomain | Format-Table -AutoSize  | Out-String
	Get-Queue -Server Exchange_HUB_Server | select DeliveryType,Status,MessageCount,NextHopDomain | Format-Table -AutoSize | Out-String
	
	foreach($MailboxServer in $MEMailBoxServers){
	
        "****Storage Group Status****" 
	    Get-StorageGroupCopyStatus -Server $MailboxServer | Sort StorageGroupName | Select StorageGroupName,SummaryCopyStatus,CCRTargetNode,CopyQueueLength | Format-Table -AutoSize | Out-String
        Write-Progress -Activity "System Checks" -CurrentOperation "Exchange Check" -Status "Running $MailboxServer Storage Group Status Check"
	    "*****Mailbox Database Status****" 
        Write-Progress -Activity "System Checks" -CurrentOperation "Exchange Check" -Status "Running $Mailboxserver Mailbox Database Status Check"
        $o = Get-ClusteredMailboxServerStatus -Identity MyMailBoxServer | select operationalmachines
		$o = $o.operationalmachines | Where-Object {$_ -like "*active*"} 
		$o = $O -split " " 

		$MachineName = [string]$o[0]
		$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $MachineName)
		$regKey= $reg.OpenSubKey("System\\CurrentControlSet\\Services\MSExchangeSA\\Parameters\\MyMailboxServerName")
		$node = [string]$regkey.GetValue("EnableOabGenOnThisNode")

		if ($MachineName -ne $node){
        "Active Node = $machineName
Offline Addressbook Generation is set to the passsive node $Node
Change the registry setting:
HKEY_LOCAL_MAHCINE\System\CurrentControlSet\Services\MSExchangeSA\Parameters\MyMailboxServerName\EnableOabGenOnThisNode
Value = $machineName " 
}
		else{
        "Active Node = $machineName
Offline Addressbook Generation is running on the active Node $MachineName. 
No Action required." 
}

   
        # Get the mailbox databases from the server

        $mbdatabases = Get-MailboxDatabase -Server $MailboxServer -status | Sort-Object -Property Name 
 

        # Get the public folder databases from the server

        $pfdatabases = Get-PublicFolderDatabase -Server $MailboxServer -status | Sort-Object -Property Name


        # Create an array for the databases

        $databases = @()
        $TotalDBSize = 0 
 

        # Check if mailbox databases were found on the server

        If ($mbdatabases) {

          # Loop through the databases

          ForEach ($mdb in $mbdatabases) {

                # Create an object to store information about the database

                $db = "" | Select-Object Name,Mounted,LastFullBackup,LastIncrementalBackup,EdbFilePath,DefragStart,DefragEnd,DefragDuration,DefragInvocations,DefragDays,Size,WhiteSpace,VolumeSize,Healthy

            

                # Populate the object

                $db.Name = $mdb.Name.ToString()
            
                $db.Mounted =[string] $mdb.Mounted

                $db.LastFullBackup =[string] (get-date([string]$mdb.LastFullBackup) -Format "yyyyMMdd")

            
                $db.LastIncrementalBackup = [string] (get-date([string]$mdb.LastIncrementalBackup) -Format "yyyyMMdd")

                $db.EdbFilePath = $mdb.EdbFilePath.ToString()
                        
                $path = ([string]$mdb.EdbFilePath.PathName) -replace "E:\\", "\\$MachineName\E$\"
            
                $dbsize = (Get-ChildItem $path).Length
	        
                $TotalDBSize += $dbsize

                $db.Size = $dbsize
            
                            
                # Add this database to the array

                $databases = $databases + $db

          } 

        }

 

        # Check if public folder databases were found on the server

        If ($pfdatabases) {

            # Loop through the databases

            ForEach ($pfdb in $pfdatabases) {

                # Create an object to store information about the database

                $db = "" | Select-Object Name,Mounted,LastFullBackup,LastIncrementalBackup,EdbFilePath,DefragStart,DefragEnd,DefragDuration,DefragInvocations,DefragDays,Size,WhiteSpace,VolumeSize,Healthy

            

                # Populate the object

                $db.Name = $pfdb.Name.ToString()

                $db.Mounted = $pfdb.Mounted

                $db.LastFullBackup = [string](get-date([string]$pfdb.LastFullBackup) -Format "yyyyMMdd")
                        
                $db.LastIncrementalBackup = [string](get-date([string]$pfdb.LastIncrementalBackup) -Format "yyyyMMdd")
            
                $db.EdbFilePath = $pfdb.EdbFilePath.ToString()
                        
                $path = ([string]$pfdb.EdbFilePath.PathName) -replace "E:\\", "\\$MachineName\E$\"
                $dbsize = (Get-ChildItem $path).Length
	        
                $TotalDBSize += $dbsize

                $db.Size = $dbsize
            
                # Add this database to the array

                $databases = $databases + $db
            } 

        }

 

        # Retrieve the events from the local Application log, filter them for ESE messages

    

        # Create an array for the output

        $out = @()
    
 

        # Loop through each of the databases and search the event logs for relevant messages
            
            $700Logs = $machines | foreach {Get-WinEvent -ComputerName $_ -FilterHashtable @{id=700;logname='application';providername ='ESE'} -MaxEvents 200 -ErrorAction SilentlyContinue}
            $701Logs = $machines | foreach {Get-WinEvent -ComputerName $_ -FilterHashtable @{id=701;logname='application';providername ='ESE'} -MaxEvents 200 -ErrorAction SilentlyContinue}
            $703Logs = $machines | foreach {Get-WinEvent -ComputerName $_ -FilterHashtable @{id=703;logname='application';providername ='ESE'} -MaxEvents 200 -ErrorAction SilentlyContinue}
            $1221Logs = $machines | foreach {Get-WinEvent -ComputerName $_ -FilterHashtable @{id=1221;logname='application';providername = 'MSExchangeIS Mailbox Store'} -MaxEvents 200 -ErrorAction SilentlyContinue}
            $logs += $700Logs + $701Logs + $703Logs + $1221Logs

        ForEach ($db in $databases) {
            
            $dbName = $db.name
            Write-Progress -Activity "System Checks" -CurrentOperation "Exchange Check" -Status "Running $dbName Status Check"

            # Create the search string to look for in the Message property of each log entry

            $s = "*" + $dbName + "*"

           

            # Search for an event 701 or 703, meaning that online defragmentation finished

            $end = $logs | where { $_.Message -like "$s" -and ($_.Id -eq 701 -or $_.Id -eq 703)} | select-object -First 1
            $endTime =[datetime] $end.TimeCreated
       
            # Search for the first event 700 which preceeds the finished event

            $start = $logs | where {$_.Message -like "$s" -and $_.Id -eq 700 -and ([datetime]$_.TimeCreated) -le $endTime} | select-object -First 1

            # Make sure we found both a start and an end message

            $WhiteSpace = $logs | where {$_.Message -like "$s" -and $_.Id -eq 1221 } | select-object -First 1

      

            # Get the start and end times

            $db.DefragStart =[string] (Get-Date($start.TimeCreated) -Format "yyyyMMdd")

            $db.DefragEnd =[string] (Get-Date($end.TimeCreated) -Format "yyyyMMdd")

 

            # Parse the end event message for the number of seconds defragmentation ran for

            $WhiteSpace.Message -match "has .* megabytes" >$null

            $numWhiteSpace = [float]($Matches[0].Split(" ")[1]) * 1MB
            if($numWhiteSpace -ge (100 * 1MB)){
                $db.WhiteSpace= ([math]::Round(($numWhiteSpace / 1gb),2)).Tostring() + " GB"
            }
            else{
                $db.WhiteSpace= ([math]::Round(($numWhiteSpace / 1mb),2)).Tostring() + " MB"
            }
            $NumSize = [float]$db.Size  

            $db.Size = ([math]::Round( ($db.size / 1GB),2)).ToString() + " GB" 
                
            $db.DefragDuration = [string] (([math]::round(([float]([timespan]($end.TimeCreated - $start.TimeCreated)).totalhours), 2)).ToString() + " Hrs")
        
            # Parse the end event message for the number of invocations and days

            $end.Message -match "requiring .* invocations over .* days" >$null

            $db.DefragInvocations =[string] ($Matches[0].Split(" ")[1])

            $db.DefragDays =[string] ($Matches[0].Split(" ")[4])
        
            $db.VolumeSize = 300 * 1GB


            $perFree = ($numWhiteSpace / $numSize) * 100
       

            if( $perFree -le 20 -and $db.DefragInvocations  -and $numSize -le ($db.volumeSize * .8)) {$db.healthy  = $true}
            else{$db.healthy = $false}

      
      

          # Add the data for this database to the output

          $out = $out + $db

        }

 

    # Print the output
    Write-Progress -Activity "System Checks" -CurrentOperation "Exchange Check" -Status "Completed $Mailboxserver Mailbox Database Status Check"
    $out |select Name,Mounted,LastFullBackup,LastIncrementalBackup,DefragStart,DefragEnd,DefragDuration,Size,WhiteSpace,Healthy| Format-Table -Property * -autosize | Out-String -Width 4096
    $TotalDBSize = [math]::Round(($TotalDBSize/ 1GB),2)
    "Total Database size = $TotalDBSize GB
    `n`n"
    } 
}

#
#    File & Print Checks
# 
Function FileandPrintCheck{
    $sharearray=@(
    '\\DFSServer\ShareName01'
    '\\DFSServer\ShareName02'
    '\\DFSServer\ShareName03'
)
	"****Checking Print/Share server status****" 
	$PFShare = @("PrintServer01","PrintServer02","PrintServer03","FileServer01","FileServer02") | ForEach-Object { 
    $Computer = $_ 
    try { 
       Get-WmiObject -class Win32_ConnectionShare  -namespace root\cimv2 -comp $Computer -EA Stop |  
      Group-Object Antecedent | 
      Select-Object @{Name="ComputerName";Expression={$Computer}}, 
                    @{Name="Share"       ;Expression={(($_.Name -split "=") |  
                    Select-Object -Index 1).trim('"')}}, 
                    @{Name="Connections" ;Expression={$_.Count}}
      
      } 
      catch  
        { 
          return  
        } 
	}
    $PFShare | Format-Table -Property * -AutoSize | Out-String -Width 4096

	foreach ($share in $sharearray) {
		if (Test-Path $share) {"Check path to $share : Successful" }
		else {"Error reaching $share , please investigate" }
	}
	
}
#
#    Web Checks
# 
function Webcheck{
    #
    #    Test Port
    # 
    function testport{
      Param([string]$srv,$port=443,$timeout=3000,[switch]$verbose)
 
      # Test-Port.ps1
      # Does a TCP connection on specified port (135 by default)
 
      $ErrorActionPreference = "SilentlyContinue"
 
      # Create TCP Client
      $tcpclient = new-Object system.Net.Sockets.TcpClient 
 
      # Tell TCP Client to connect to machine on Port
      $iar = $tcpclient.BeginConnect($srv,$port,$null,$null)
 
      # Set the wait time
      $wait = $iar.AsyncWaitHandle.WaitOne($timeout,$false)
 
      # Check to see if the connection is done
      if(!$wait)
      {
        # Close the connection and report timeout
        $tcpclient.Close()
        if($verbose){"Connection Timeout"}
        Return $false
      }
      else
      {
        # Close the connection and report the error if there is one
        $error.Clear()
        $tcpclient.EndConnect($iar) | out-Null
        if(!$?){if($verbose){$error[0]};$failed = $true}
        $tcpclient.Close()
      }
      # Return $true if connection Establish else $False
      if($failed){return $false}else{return $true}
    } 
    #
    #    Test Port
    #
    $httpsarray=@(
    'ISE_Server51'
    'CiscoWorksServer'
    'ServiceDeskURL'
    'OWA')
	foreach ($https in $httpsarray){
		if (testport $https){"  Port 443 open to $https" }
		else {"  Error connecting to $https on port 443" }
	}
	
	
}
#
#    Run Check Jobs
# 	
function RunRegularWatchChecks{
	cls
    $output = @()
    $Date = [string](get-date)
	
    $output += $Date
	
    Write-Progress -Activity "System Checks" -CurrentOperation "Starting" -Status "Running" 
    
    #Blackberry Check
    $Output += BlackberryCheck
    Write-Progress -Activity "System Checks" -CurrentOperation "Blackberry Check" -Status "Running"
    sleep -Seconds 3
    Write-Progress -Activity "System Checks" -CurrentOperation "Blackberry Check" -Status "Completed"
    
    cls
    

    #Exchange Check
    Write-Progress -Activity "System Checks" -CurrentOperation "Exchange Check" -Status "Starting"
    Start-Sleep 3 
    cls 
	Write-Progress -Activity "System Checks" -CurrentOperation "Exchange Check" -Status "Running"
    $Output +=ExchangeCheck
    Write-Progress -Activity "System Checks" -CurrentOperation "Exchange Check" -Status "Completed"
    Start-Sleep 3 
    #File and Print Check
    Write-Progress -Activity "System Checks" -CurrentOperation "File and Print Check" -Status "Starting"
    Start-Sleep 3 
    cls 
	Write-Progress -Activity "System Checks" -CurrentOperation "File and Print Check" -Status "Running"
    $Output +=FileandPrintCheck
    Write-Progress -Activity "System Checks" -CurrentOperation "File and Print Check" -Status "Completed"
    Start-Sleep 3 
    #Web Check
    Write-Progress -Activity "System Checks" -CurrentOperation "Web Check" -Status "Starting"
    Start-Sleep 3 
    cls 
	Write-Progress -Activity "System Checks" -CurrentOperation "Web Check" -Status "Running"
    $Output +=Webcheck
    Write-Progress -Activity "System Checks" -CurrentOperation "Web Check" -Status "Completed"
    Start-Sleep 3 
    #Create Check Log
    Write-Progress -Activity "System Checks" -CurrentOperation "Creating Check Log" -Status "Running"
    Set-Content -Path 'UpdateToChecks.txt' -Value $output 
    Write-Progress -Activity "System Checks" -CurrentOperation "Creating Check Log" -Status "Complete"
    Start-Sleep -Seconds 5
	cls
    Write-Progress -Activity "System Checks" -CurrentOperation "Alert" -Status "Check your email for BlackBerry confirmation receipt."
	Sleep 5
    Write-Progress -Activity "System Checks" -Status "Complete" -Completed
    
    if ($runonreplay -eq "y"){
		
        notepad UpdateToChecks.txt
		
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		$objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon 
		#$objNotifyIcon.Icon = "d:\icon1.ico"
		$objNotifyIcon.BalloonTipIcon = "Warning" 
		$objNotifyIcon.BalloonTipText =@"
		Please copy text from notepad and post to 
		
		Here was the patch of the NASO 24Hr Watch Log Path
				
"@
		
		$objNotifyIcon.BalloonTipTitle = " Time to Update the Watch Log ! "
		$objNotifyIcon.Visible = $True 
		$objNotifyIcon.ShowBalloonTip(10000)
        
        $StartTime = [datetime]::Now.AddHours(2) 
		Start-Job -ScriptBlock {Start-Sleep -Seconds 7200} -Name 'replay' | out-null
        
        $ck = Get-Job -Name "replay"
        cls 
        $jobState = $ck.State.ToString()
        while($jobState -ne "Completed"){
            $timeNow = get-date
            $timeRemaining = ($startTime - $timeNow).totalSeconds
            Write-Progress -Activity "System Checks" -CurrentOperation "Run next check" -Status "Waiting" -SecondsRemaining $timeRemaining 
        }       
		Remove-Job -Name 'replay' -Force | Out-Null
		RunRegularWatchChecks
	}
    else{ notepad UpdateToChecks.txt}

	
}
#
#    Exchange Checks Query
#
function RegularWatchChecksQuery{
    if ($runonreplay -ne "y"){
        cls
        Write-Host @"
Regular Watch Checks     
--------------------------------------------------------
If you choose to repeat, it will run,
wait 2 hours, and run again, until the end of time    
--------------------------------------------------------
"@
    $runonreplay = read-host "Repeat on a 2 hour cycle? [y/n]"
    
    }
    
	RunRegularWatchChecks	
}
#
#  VMware PowerCLI window.  Should learn how to open one that runs the get-vm command correctly, or get some perf-stats
#
function openvmware {
    $vmwarepowercli=". 'C:\Program Files\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1'"
    Write-Host "    |"
    cmd /c start powershell -psc "C:\Program Files\VMware\Infrastructure\vSphere PowerCLI\vim.psc1" -noe -c $vmwarepowercli
    Write-Host @"          
    |   If you have VMware PowerCLI installed - 
    |
    |   It should have auto-connected you:
    |
    |   type 'get-vm' for status and resource allotment.
    |
"@
}
#
#    Body of the Script
#
do{    
    Set-Color 20
    Clear-Host
    Write-Host @"
    |
    |          $currentversion
    |
    |   - NIPR System watch Toolbox -
    |
    |            pick your poison...
    |
    |  [0]  Help menu / Read-me
    |
    | *[1]  Regular Watch Checks
    |
    |  [2]  Defrag and Analyze
    |
    |  [3]  Disk Cleanup Sequence
    |
    |  [4]  Monitor Svr Connectivity
    |
    |  [5]  Initiate Server Reboots
    |
    | *[6]  Post-Reboot Service Checks
    |
    |  [7]  VMware PowerCLI window
    |
    |  [8]  Storage Utilization Report
    |
    |  [9]  Backup Exec Logs
    |
    |  [10] Exit
    |
"@
    $poison = read-host "    |   Make your choice [0-10]"
    write-host "    |"
    Set-Color 2F
    If ($poison -eq 0){helpmenu}
      elseIf ($poison -eq 1){RegularWatchChecksquery}
        elseIf ($poison -eq 2){defragglerock}
          elseIf ($poison -eq 3){callthecleaner}
            elseIf ($poison -eq 4){start-monitor}
              elseIf ($poison -eq 5){rebootserver}
                elseIf ($poison -eq 6){autostopped}
                  elseIf ($poison -eq 7){openvmware}
                    elseIf ($poison -eq 8){utilizationreport}
                      elseIf ($poison -eq 9){CommVaultReport}
                        elseIf ($poison -eq 10){exitscript}
    Write-Host "`n"
    $round = read-host "`tBack to main menu? [y/n]"
    if ($round -ne 'y') {exitscript}
    
}while ($round -eq "y")
