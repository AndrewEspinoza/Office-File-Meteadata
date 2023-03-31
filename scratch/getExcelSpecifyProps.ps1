function Get-ExcelMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath,
        
        [string[]]$Extensions = @("*.xls", "*.xlsx", "*.xlsm"),
        
        [string]$OutputPath
    )
    
    # Get all Excel files in the folder
    $excelFiles = Get-ChildItem $FolderPath -Include $Extensions -Recurse
    
    # Start Excel and loop through the Excel files
    $excel = New-Object -ComObject Excel.Application
    $results = @()
    foreach ($file in $excelFiles) {
        $workbook = $excel.Workbooks.Open($file.FullName)
        $fullName = $workbook.FullName
        $author = $workbook.Author
        $results += [PSCustomObject]@{
            Filename = $fullName
            Author = $author
        }
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }

    # Quit Excel and release the associated resources
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel
    
    # Write the results to a CSV file if specified
    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation
    }
    
    # Return the results as an array of objects
    return $results
}
