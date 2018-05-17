<#
.SYNOPSIS
  Reads Spreadsheet Data
  
.DESCRIPTION
  Opens Excel spreadsheet
  Copies data to custom object array
  Copies array to hashtable
  
.NOTES
  File Name: Read-Spreadsheet
  Author: Jay Berkovitz

.REQUIREMENTS
  PS Version 5
  Import XLSx file
#>

# -- This Function converts xlsx to csv for import -- #
function Import-Xls 
{ 
    [CmdletBinding(SupportsShouldProcess=$true)] 
    Param( 
        [parameter( 
            mandatory=$true,  
            position=1,  
            ValueFromPipeline=$true,  
            ValueFromPipelineByPropertyName=$true)] 
        [String[]] 
        $Path, 
        [parameter(mandatory=$false)] 
        $Worksheet = 1, 
        [parameter(mandatory=$false)] 
        [switch] 
        $Force 
    ) 
    Begin 
    { 
        function GetTempFileName($extension) 
        { 
            $temp = [io.path]::GetTempFileName(); 
            $params = @{ 
                Path = $temp; 
                Destination = $temp + $extension; 
                Confirm = $false; 
                Verbose = $VerbosePreference; 
            } 
            Move-Item @params; 
            $temp += $extension; 
            return $temp; 
        } 
        $xlFileFormats = @{ 
            '.csv'  = 6;        # 6, 22, 23, 24 
            '.xls'  = -4143;    # -4143, 1, 16, 18, 29, 33, 39, 43 
            '.xlsx' = 51;       # 
            '.xml'  = 46;       # 
        } 
        $xl = New-Object -ComObject Excel.Application; 
        $xl.DisplayAlerts = $false; 
        $xl.Visible = $false; 
    } 
    Process 
    { 
        $Path | ForEach-Object { 
            if ($Force -or $psCmdlet.ShouldProcess($_)) { 
                $fileExist = Test-Path $_ 
                if (-not $fileExist) { 
                    Write-Error "Error: $_ does not exist" -Category ResourceUnavailable;             
                } else { 
                    # create temporary .csv file from excel file and import .csv 
                    # 
                    $_ = (Resolve-Path $_).toString(); 
                    $wb = $xl.Workbooks.Add($_); 
                    if ($?) { 
                        $csvTemp = GetTempFileName(".csv"); 
                        $ws = $wb.Worksheets.Item($Worksheet); 
                        $ws.SaveAs($csvTemp, $xlFileFormats[".csv"]); 
                        $wb.Close($false); 
                        Remove-Variable -Name ('ws', 'wb') -Confirm:$false; 
                        Import-Csv $csvTemp; 
                        Remove-Item $csvTemp -Confirm:$false -Verbose:$VerbosePreference; 
                    } 
            }} 
    }} 
    End 
    {$xl.Quit(); 
    Remove-Variable -name xl -Confirm:$false; 
    [gc]::Collect();} 
} 


# -- Import Spreadsheet with Function and chop off the header -- #
$sheets = ".\spreadsheet.xlsx" | Import-Xls -Worksheet 1
$header = ($sheets | Get-Member | Where Membertype -eq NoteProperty).Name
$string = $sheets.$header
$keys = @()

# -- Draws out a list of the keys we'll need -- #
0..$string.count | Foreach{If (!$string[$_]){$keys += $string[$_+1]}}

# -- Groups the settings together in PSobjects with thier key as "paragraphs" -- #
$count = 0
$paragraphs = @()
1..$string.count | Foreach{
    $currObj = $string[$_]
    If ($currObj -and $keys -notcontains $currObj){
        $paragraph = New-Object -TypeName PSObject -Property (@{
            'key' = $keys[$count]
            'val' = $currObj.replace('-- ','')
        }) 
    $paragraphs += $paragraph
    }
    If (!$currObj){$count++}
}

# -- Generates a hashtable from $paragraphs, making keys unique -- #
$keys = @()
$hashtable = @{}
$paragraphs | Foreach {
    $key = $_.key
    $val = "$_.val"
    If ($keys -notcontains $key){
          $hashtable.Add($key,$_.val)
    }Else{$hashtable.$key += "`n"+$_.val}
    $keys += $key
}

# -- Outputs the $paragraphs to a CSV file -- #
$SaveDate = [string](Get-Date -Format yyyyMMdd_HHmm)
$output = "c:\a\"+$Savedate+".csv"
$paragraphs | export-csv -NoTypeInformation $output

# -- Displays data in original format
$hashtable.Keys.ForEach({"`n"+$_+"`n"+$hashtable.$_})
