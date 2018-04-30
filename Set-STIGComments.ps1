$user = "Jay Berkovitz WSMIS ISSO"
$initpath = "c:\dwn\"

# This function creates the dialog box to choose the checklist file

Function Get-FileName($initialDirectory)
{   
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.ShowHelp = $true
    $OpenFileDialog.filter = "All files (*.ckl)| *.ckl"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
} #end function Get-FileName

#Sets the $path variable to the file you chose in the dialog box
$path = Get-FileName -initialDirectory $initpath

#Loads the contents as XML
$xml = [xml](Get-Content $path)


#sets the $date variable to the current date and formats it as dd/mm/YYYY and formats the savedate to ISO yyyyMMdd
$date = Get-Date -format d
$savedate = (Get-Date).tostring("yyyyMMdd")

<#
For each $Attr (node) at the VULN level of the tree, 
check the STATUS node for a match to "NotAFinding" 
and then set the COMMENTS node to "Reviewed by Username on dd/mm/YYYY"
#>
ForEach ($Attr in $xml.CHECKLIST.STIGS.iSTIG.VULN) {
    If ($Attr.STATUS -match "NotAFinding") {
        $Attr.COMMENTS = "Reviewed by $user on $date"
    }
}

#Save the now modified xml back to the file you initially loaded.
$destination = Split-Path -Path $path -Parent
$filename = [io.path]::GetFileNameWithoutExtension("$path")
$xml.Save($destination + "\" + $filename + "_$savedate.ckl")