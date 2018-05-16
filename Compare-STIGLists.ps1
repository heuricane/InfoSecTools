<#
.SYNOPSIS
  Pulls the STIGs on the short list from the long list
  
.NOTES
  File Name: Compare-STIGLists.ps1
  Author: Jay Berkovitz
  Link: github.com/heuricane/InfoSecTools/blob/master/Compare-STIGLists.ps1
  
.REQUIREMENTS
  PS Version 5
  Two checklists input
#>

# -- Put the two lists in C:\tmp -- #
Set-Location C:\tmp

# -- Import the two lists to $vars -- #
$long  = Import-CSV -Path longlist.csv
$short = Import-CSV -Path shortlist.csv

# -- Compare and output -- #
$list = $long | Where-Object {$short.STIG_ID -contains $_.STIG_ID}
$list | Export-CSV -path list.csv -NoTypeInformation
