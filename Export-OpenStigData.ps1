<#
.SYNOPSIS
    Saves a CSV containing information on all Open stigs
.DESCRIPTION
    Saves a CSV containing information on all Open stigs
.PARAMETER CKLDirectory
    Full path to a directory containing your checklists
  
.EXAMPLE
    .\Export-OpenStigData.ps1 -CKLDirectory 'C:\dwn\' -SavePath 'C:\tmp\OpenChecks.csv'
#>
Param([Parameter(Mandatory=$true)]$CKLDirectory,[Parameter(Mandatory=$true)]$SavePath)
#Check if module imported
if ((Get-Module | Where-Object -FilterScript {$_.Name -eq "StigSupport"}).Count -le 0)
{
    #End if not
    Write-Error "Please import StigSupport.psm1 before running this script"
    return
}

#List all CKL Files
$CKLs = Get-ChildItem -Path $CKLDirectory -Filter "*.ckl"

#Initialize Results
$STIGData = @()

Foreach ($CKL in $CKLs)
{
    #Load this CKL
    $CKLData = Import-StigCKL -Path $CKL.FullName
    $HostData = Get-CKLHostData -CKLData $CKLData
    #Grab data on all stigs.
    #Format of @{Status,Finding,Comments,VulnID,Comments}
    $Stigs = Get-VulnCheckResult -XMLData $CKLData
    $Asset = $HostData.HostName
    foreach ($Stig in $Stigs)
    {
        #Requested STIG, Vuln ID, Rule Title/Name, Check Content, Fix Text, Comments
        if ($Stig.Status -notmatch "NotAFinding") {
            Write-Host "Need $($Stig.VulnID)"
            if (($STIGData | Where-Object -FilterScript {$_.VulnID -eq $Stig.VulnID})) {
                ($STIGData | Where-Object -FilterScript {$_.VulnID -eq $Stig.VulnID}).Count++;
            } else {
                $ToAdd = New-Object -TypeName PSObject -Property @{VulnID=$Stig.VulnID;RuleTitle="";CheckContent="";FixText="";Comments="";Asset=""}
                $ToAdd.RuleTitle = Get-VulnInfoAttribute -XMLData $CKLData -VulnID $Stig.VulnID -Attribute "Rule_Title"
                $ToAdd.CheckContent = Get-VulnInfoAttribute -XMLData $CKLData -VulnID $Stig.VulnID -Attribute "Check_Content"
                $ToAdd.FixText = Get-VulnInfoAttribute -XMLData $CKLData -VulnID $Stig.VulnID -Attribute "Fix_Text"
                $ToAdd.Comments = Get-VulnInfoAttribute -XMLData $CKLData -VulnID $Stig.VulnID -Attribute "Comment"
                $ToAdd.Asset = $Asset
                $STIGData += $ToAdd
            }
        } else {
            Write-Host "Skip $($Stig.VulnID)"
        }
    }
}

$STIGData | Export-Csv -Path $SavePath -NoTypeInformation