$rootFolder = "C:\temp\Files\*"       # * for wildcard is required. The top level folder of the file structure where the PDFs are.
$OutputFile = "C:\temp\output.csv"    # The folder and name of the CSV for your output.


$files = (Get-ChildItem -Path $rootFolder -Recurse)

$output = @()

foreach ($file in $files) {

$sha1 = (Get-FileHash -Path $file -Algorithm SHA1).Hash
$md5 = (Get-FileHash -Path $file -Algorithm MD5).Hash

$row = New-Object PSObject
$row | Add-Member -MemberType NoteProperty -Name "FileName" -Value $file.Name
$row | Add-Member -MemberType NoteProperty -Name "SHA1Hash" -Value $sha1
$row | Add-Member -MemberType NoteProperty -Name "MD5Hash" -Value $md5

$output += $row

}

$output | Export-Csv -Path $OutputFile -NoTypeInformation