$OutputFile = "C:\temp\output.csv"
$files = (Get-ChildItem -Path C:\temp\FIles\* -Recurse)

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