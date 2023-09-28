# Specify the path to the folder and the search value
$folder_path = "C:\Users\petartr\OneDrive - Microsoft\Desktop\Gov.SI\"
$file_extension = "*.rpt"
$search_values = @("[INFORMATION", "[WARNING" )
$excluded_values = @("Discontinued", "Deprecated" )
$Results_file = $folder_path + "results.csv"



$results = @()


# Get all the text files in the folder
$file_paths = Get-ChildItem -Path $folder_path -Filter $file_extension -Recurse
   


# Loop through each file
foreach ($file_path in $file_paths) {
    # Open the file
    Write-Host $file_path.FullName  -ForegroundColor Yellow
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
                $CategorySplit =  $line.Substring(0,$start_pos).Split(" ")


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
                    #Text = $line.Substring($end_pos+1,$line.Length-($end_pos+1) )
                    Row = $line

                    }
                }
            }
        }

    }
}



#$results | Export-Csv -Path "C:\Users\petartr\Downloads\BP_Check_Files\results.csv" -NoTypeInformation
$results | Export-Csv -Path $Results_file -NoTypeInformation
 