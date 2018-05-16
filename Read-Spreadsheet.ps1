<#
.SYNOPSIS
  Reads Spreadsheet Data
  
.DESCRIPTION
  Opens Excel spreadsheet
  Copies data to table
  
.NOTES
  File Name: Read-Spreadsheet.ps1
  Author: Jay Berkovitz

.REQUIREMENTS
  PS Version 5
  Input STIG XML file
#>

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
            # single worksheet formats 
            '.csv'  = 6;        # 6, 22, 23, 24 
            '.dbf'  = 11;       # 7, 8, 11 
            '.dif'  = 9;        #  
            '.prn'  = 36;       #  
            '.slk'  = 2;        # 2, 10 
            '.wk1'  = 31;       # 5, 30, 31 
            '.wk3'  = 32;       # 15, 32 
            '.wk4'  = 38;       #  
            '.wks'  = 4;        #  
            '.xlw'  = 35;       #  
             
            # multiple worksheet formats 
            '.xls'  = -4143;    # -4143, 1, 16, 18, 29, 33, 39, 43 
            '.xlsb' = 50;       # 
            '.xlsm' = 52;       # 
            '.xlsx' = 51;       # 
            '.xml'  = 46;       # 
            '.ods'  = 60;       # 
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
                } 
            } 
        } 
    } 
    End 
    { 
        $xl.Quit(); 
        Remove-Variable -name xl -Confirm:$false; 
        [gc]::Collect(); 
    } 
} 

$sheets = ".\spreadsheet.xlsx" | Import-Xls -Worksheet 1
$header = ($sheets | Get-Member | Where Membertype -eq NoteProperty).Name
$string = $sheets.$header
$hashtable = @{}
$paragraphs = @()
$keys = @()
$count = 0

0..$string.count | Foreach{If (!$string[$_]){$keys += $string[$_+1]}}

$count = 0
$paragraphs = @()
1..$string.count | Foreach{$currObj = $string[$_]
    If ($currObj -and $keys -notcontains $currObj){
        $paragraph = New-Object -TypeName PSObject -Property (@{
            'key' = $keys[$count]
            'val' = $currObj.replace('-- ','')
        }) 
    $paragraphs += $paragraph
    }
    If (!$currObj){$count++}
}

$paragraphs | export-csv -NoTypeInformation c:\a\_test.csv
