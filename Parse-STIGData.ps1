<#
.SYNOPSIS
  Inputs XML > Outputs Spreadsheet
  
.DESCRIPTION
  Opens STIG XML data (or a checklist)
  Searches for Elements that match criteria
  Cleans/Exports data to a CSV file
  
.NOTES
  File Name: Parse-STIGData.ps1
  Author: Jay Berkovitz
  Link: github.com/heuricane/InfoSecTools/blob/master/Parse-STIGData.ps1
  
.REQUIREMENTS
  PS Version 5
  Input STIG XML file
#>

# -- Preset the Variables -- #
$path = "C:\a\Checklist.xml"
$search = "Status"
$find = "Open"
$custList=@()

# -- Loads the contents as XML and searches -- #
$xml = [xml](Get-Content $path)
$list = $xml.CHECKLIST.STIGS.iSTIG.VULN | Where $search -eq $find

# -- Loop Through Findings -- #
0..($list.Count-1) | Foreach {

# -- Truncate Comments that Exceed the CSV Max Character Length -- #
    $comment = $list[$_].COMMENTS.Trim()
    If($comment.Length -gt 32767){$comment = $comment.Substring(0,32756)}

# -- Apply settings to object for Export -- #
    $customObject = New-Object -TypeName PSObject -Property (@{
        'VULN_IDs' = $list[$_].STIG_DATA.ATTRIBUTE_DATA[0].Trim();
        'Statuses' = $list[$_].STATUS.Trim();
        'Detailed' = $list[$_].FINDING_DETAILS.Trim();
        'Comments' = $comment
        })
    $custList += $customObject
}
# -- Export -- #
$saveDate = (Get-Date).tostring("yyyyMMdd")
$output = "C:\a\"  +$saveDate + "_Custom.csv"
$custList | Select VULN_IDs,Statuses,Detailed,Comments | 
    Export-CSV -Path $output -NoTypeInformation
