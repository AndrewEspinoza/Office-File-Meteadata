function Get-DocumentProperties {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$FolderPath,
        [Parameter(Mandatory=$true, Position=1)]
        [string]$OutputFile
    )

    # Get all Word files in the folder
    $wordFiles = Get-ChildItem $FolderPath -Include *.doc,*.docx,*.docm -Recurse

    # Create an array to hold the output
    $output = @()

    # Loop through each document and extract all available properties
    foreach ($document in $wordFiles) {
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($document.FullName)

        $properties = $doc | Get-Member -MemberType Property

        $row = New-Object PSObject
        $row | Add-Member -MemberType NoteProperty -Name "FileName" -Value $document.Name
        foreach ($property in $properties) {
            $value = $doc.$($property.Name)
            $row | Add-Member -MemberType NoteProperty -Name $property.Name -Value $value
        }
        $output += $row

        $doc.Close()
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }

    # Export the output to a CSV file
    $output | Export-Csv -Path $outputFile -NoTypeInformation
}
Get-DocumentProperties -folderPath "C:\temp\" -outputFile "C:\temp\word_out.csv"