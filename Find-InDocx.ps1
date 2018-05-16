<#
.SYNOPSIS
  Searches Doc for words in (Parentheses)
  
.DESCRIPTION
    Opens the file, searches for words, filters, outputs
  
.NOTES
  File Name: Find-InDocx.ps1
  Author: Jay Berkovitz
  Link: github.com/heuricane/InfoSecTools/blob/master/Find-InDocx.ps1
  
.REQUIREMENTS
  PS Version 5
  Input Word Doc
#>

$list = @()
$output = @()

# -- Target the file -- #
$filePath = "C:\tmp\file.docx"
$doc = Get-ChildItem -path $filePath

# -- Open the File -- #
$application = New-Object -comobject word.application 
$document = $application.documents.open("$doc", $false, $true)
$application.visible = $False
$matchCase = $false 
$matchWholeWord = $false 
$matchWildCards = $true 
$matchSoundsLike = $false 
$matchAllWordForms = $false 
$forward = $true 
$wrap = 1
$range = $document.content
$null = $range.movestart()
$string = $document.Content.Text

# -- Search for all words for words in Parentheses -- #
$count = ($string | measure-object -word).Words
1..$count | Foreach {
    '(' + $String.Split('()')[$_] + ')'
    $list += $String.Split('()')[$_]
}

# -- Remove blank entries and sentences in Parenthesis -- #
Foreach ($line in $list){
    If ($line -ne ""){
        If($line.length -lt 16){
            $output += $line
        }
    }
}
# -- Saves the Output -- #
$save = (Get-Date).tostring("yyyyMMddhhmmss")
$outputPath = "C:\tmp\list-" +$save+ ".csv"
$output | Sort -Unique > $outputPath

$document.close()
$application.quit()
