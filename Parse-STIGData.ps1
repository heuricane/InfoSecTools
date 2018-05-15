<#
.SYNOPSIS
  Opens STIG XML data (or a checklist)
  Searches for Elements that match criteria
  Cleans/Exports data to a CSV file
#>

# Preset Variables
$path = "C:\a\Checklist.xml"
$savedate = (Get-Date).tostring("yyyyMMdd")
$Output = "C:\a\"  +$savedate + "_Custom.csv"
$search = "Status"
$find = "Open"

#Loads the contents as XML and searches
$xml = [xml](Get-Content $path)
$Vulns = $xml.CHECKLIST.STIGS.iSTIG.VULN
$list = $vulns | where $search -eq $find
$max = $list.count - 1
$CustList = @()
$count = 0

# Loop Through Findings
Do{

# Truncate Comments that Exceed the Max Character Length
$Comment = $list[$count].COMMENTS.Trim()
If($Comment.Length -gt 32767){$Comment = $Comment.Substring(0,32756)}

# Apply settings to object for Export
    $CustomObject = New-Object -TypeName PSObject -Property (@{
        'VULN_IDs' = $list[$count].STIG_DATA.ATTRIBUTE_DATA[0].Trim();
        'Statuses' = $list[$count].STATUS.Trim();
        'Detailed' = $list[$count].FINDING_DETAILS.Trim();
        'Comments' = $Comment
        })
    $CustList += $CustomObject
    $count++
}While ($count -le $max)

# Export
$CustList | Export-CSV -NoClobber -NoTypeInfo -Path $Output
