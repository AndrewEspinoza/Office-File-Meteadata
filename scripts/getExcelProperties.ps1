function Get-ExcelProperties {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$FolderPath,
        [Parameter(Mandatory=$true, Position=1)]
        [string]$OutputFile
    )

    # Get all Excel files in the folder
    $excelFiles = Get-ChildItem $FolderPath -Include *.xls,*.xlsx,*.xlsm -Recurse

    # Create an array to hold the output
    $output = @()

    # Loop through each Excel file and extract the author property
    foreach ($file in $excelFiles) {
        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Open($file.FullName)

        $properties = $workbook | Get-Member -MemberType Property

        $row = New-Object PSObject
        $row | Add-Member -MemberType NoteProperty -Name "FileName" -Value $file.Name
        foreach ($property in $properties) {
            $value = $workbook.$($property.Name)
            $row | Add-Member -MemberType NoteProperty -Name $property.Name -Value $value
        }
        $output += $row

        $workbook.Close($false)
        try {
            $excel.Quit()
        }catch{
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    # Export the output to a CSV file
    $output | Export-Csv -Path $OutputFile -NoTypeInformation
}
