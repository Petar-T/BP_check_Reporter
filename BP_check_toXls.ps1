﻿function Cvs2Xls{
    param (
    [string] $FolderName,
    [string] $TimeStamp,
    [string] $OutputFolder,
    [bool] $CompressContents = $true
    ) 


cd $FolderName;

$csvs = Get-ChildItem .\* -Include *.csv | Sort-Object Name
$y=$csvs.Count

    $outputfilename =  "BP_Check_output_" + $TimeStamp + ".xlsx"

    Write-Host "Creating: $outputfilename, total: $y tabs" -ForegroundColor Yellow

    $excelapp = new-object -comobject Excel.Application
    $excelapp.sheetsInNewWorkbook = $csvs.Count
    $xlsx = $excelapp.Workbooks.Add()
    $sheet=1

    foreach ($csv in $csvs)
    {
        $row=1
        $column=1
        $worksheet = $xlsx.Worksheets.Item($sheet)
        $worksheet.Name = $csv.Name
        $file = (Get-Content $csv)
        foreach($line in $file)
        {
            $linecontents=$line -split ',(?!\s*\w+")'
            foreach($cell in $linecontents)
            {
                $cell=$cell.TrimStart('"')
                $cell=$cell.TrimEnd('"')
                $worksheet.Cells.Item($row,$column) = $cell
                $column++
            }
        $column=1
        $row++
        }
    $sheet++
    }

    #$output = $FolderName + "\" + $outputfilename
    $output = "$OutputFolder\$outputfilename" 


    
    $xlsx.SaveAs($output)
    $excelapp.quit()
     
    if ($CompressContents)
    {
     Write-Host "Zipping individual files into archive $OutputFolder\CvsArchive.zip" -ForegroundColor Yellow
     foreach ($csv in $csvs)
     {
        Compress-Archive -Path $csv -DestinationPath "$OutputFolder\CvsArchive.zip" -Update
        Remove-Item -Path $csv
     }
    
    }
}

function Get-ResultsFile{
    param (
    [string] $folder_path
    ) 

$search_values = @("[INFORMATION", "[WARNING" )
$excluded_values = @("Discontinued", "Deprecated" )
#$ParentFolder = Split-Path -Parent $folder_path
#$Results_file = $ParentFolder + "\_results.csv"
$Results_file = $folder_path + "\_results.csv"

Write-Host Analyzing : $folder_path  -ForegroundColor Yellow

$results = @()

$file_paths = Get-ChildItem -Path $folder_path -Filter *.csv 
   


# Loop through each file
foreach ($file_path in $file_paths) {
    # Open the file
    #Write-Host $file_path.FullName  -ForegroundColor Yellow
    $file_content = Get-Content $file_path.FullName

    # Loop through each line in the file
    foreach ($line in $file_content) {
        foreach ($search_value in $search_values ) {
            # Check if the line contains the search value
            if ($line.Contains($search_value)) {
                # If the line contains the search value, output it to the console
                #Write-Host $file_path : [Info] $line 
                

                $start_pos = $line.IndexOf($search_value)
                $end_pos   = $line.IndexOf("]", $start_pos + $search_value.Length)
                $subRes    = $line.Substring($start_pos, $end_pos - $start_pos + 1)
                #$Category  = $line.Substring(0,$start_pos)

                $Category  = $line.Substring(0,$line.IndexOf(" ")) 
                $CategorySplit =  $line.Substring(0,$start_pos).Split(",")


                $exclusion_Found = 0
                foreach ($exc in $excluded_values){
                if ($subRes.Contains($exc)) {
                $exclusion_Found ++    }}
                  
                if ($exclusion_Found -eq 0){

                $results += [PSCustomObject]@{
                    FileName = $file_path.Name
                    SearchType = $search_value.Substring(1, $search_value.Length-1 )
                    Category = $CategorySplit[0]
                    SubCategory = $CategorySplit[1]
                    SearchedText = $SubRes
                    Row = $line
                    #Cat = $Category

                    }
                }
            }
        }

    }
}

$results | Export-Csv -Path $Results_file -NoTypeInformation
Write-Host Writing Results file to $Results_file  -ForegroundColor Yellow
} 

# thanks to https://vladdba.com/2023/01/17/save-execution-plan-files-powershell/
function Format-XML {
 [CmdletBinding()]
 Param (
        [Parameter(
            ValueFromPipeline=$true,
            Mandatory=$true)]
            [string]$XMLInput
            )
 $XMLDoc = New-Object -TypeName System.Xml.XmlDocument
 $XMLDoc.LoadXml($XMLInput)
 $SW = New-Object System.IO.StringWriter
 $Writer = New-Object System.Xml.XmlTextwriter($SW)
 $Writer.Formatting = [System.XML.Formatting]::Indented
 $XMLDoc.WriteContentTo($Writer)
 $SW.ToString()
  }

$dataSource = 'localhost'
$database = "MSDB"
#$sqlcommand = "exec dbo.usp_bpcheck @diskfrag=0" #bypassing Pshell elevation error
$sqlcommand = "exec dbo.usp_bpcheck" 

$planHeader = "Category
Check
query_plan"
$ExtractPlans=$true

$TimeGenerated=Get-Date -Format "ddMMyy_HHmm"
$Fldr="Bp_check_" + $($TimeGenerated)
$WorkingFolder = New-Item -Path  $env:TEMP -ItemType Directory -Name $Fldr

Write-Host "creating individual .cvs files in $WorkingFolder" -ForegroundColor Yellow

try{
    $connectionString = "Data Source=$dataSource; Integrated Security=True;Application Name=BP_Check PS runner ; Initial Catalog=$database"
    $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
    $command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection) 
    $command.CommandTimeout = 3600

    #Async handler of InfoMessage
    $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {param($sender, $event) if ($event.Errors[0].Class -le 10) {Write-Host "Info_Msg :  $event" -ForegroundColor Green} else {Write-Host "Error_Msg :  $event" -ForegroundColor Red} }; 
    $connection.add_InfoMessage($handler); 


    $connection.FireInfoMessageEventOnUserErrors = $true;

    $connection.Open()
    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command 
    $adapter.ContinueUpdateOnError=$true
    $dataset = New-Object System.Data.DataSet


        Write-Host "Running bp_check NOW, be patient! ..." -ForegroundColor Yellow

    $StartTime = $(get-date)
        $adapter.Fill($dataSet) | Out-Null -ErrorAction Continue
        #$connection.Close()
    $elapsedTime = NEW-TIMESPAN -Start $StartTime -End $(get-date)
    $elapsedTimeStr = '{0:00}:{1:00}' -f $elapsedTime.Minutes, $elapsedTime.Seconds
    Write-Host $("Elapsed Time : [$elapsedTimeSt]") -ForegroundColor White

    }
catch{
     Write-Host $_.Exception.Message -ForegroundColor Red
    }
Finally
{
     $connection.Close()
}

$TableNum = 1

<# code without plan extraction
ForEach($table in $dataset.Tables)
{
        $StrTableNum= $TableNum.ToString("000")
        $table | Export-Csv -Path "$WorkingFolder\Test$StrTableNum.csv" -NoTypeInformation

        $TableNum ++
        #Write-Host "Test$StrTableNum.csv"
}#>
ForEach($table in $dataset.Tables)
{
        $StrTableNum= $TableNum.ToString("000")

        $headerRow = $table.Rows[0].psobject.properties.name 
        $headerRowText =  Out-String -InputObject $headerRow
                
        if (($headerRowText.Length  -gt 0 ) -and ($ExtractPlans -eq $true))
        {
                if ($headerRowText.Substring(0,27) -eq $planHeader )
                {
                    Write-Host "Table with Plan found! [$StrTableNum]" -ForegroundColor Yellow
                        
                        $TablefolderPath = "$WorkingFolder\Table$StrTableNum"  
                        if (!(Test-Path $TablefolderPath -PathType Container)) 
                            {
                                New-Item -ItemType Directory -Force -Path $TablefolderPath | Out-Null
                            }

                    [int]$RowNum = 0
                            
                    foreach($row in $table)
                        {
                            [string]$SQLPlanFile = "Plan_$RowNum.sqlplan" 
                            $table.Rows[$RowNum]["query_plan"] | Format-XML | Set-Content  -Path "$WorkingFolder\Table$StrTableNum\$($SQLPlanFile)" -Force
                            $RowNum+=1
                        }


                } 
        }

        $table | Export-Csv -Path "$WorkingFolder\Test$StrTableNum.csv" -NoTypeInformation

        $TableNum ++
        #Write-Host "Test$StrTableNum.csv"
}


Get-ResultsFile -folder_path $WorkingFolder

Cvs2Xls -FolderName $WorkingFolder -TimeStamp $TimeGenerated -OutputFolder $WorkingFolder -CompressContents $true

ii .