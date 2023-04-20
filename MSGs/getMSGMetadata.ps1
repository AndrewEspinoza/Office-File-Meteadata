$msgRootFolder = "C:\temp"              # The top level folder of the file structure where the PDFs are.
$outputFile = "C:\My\Folder\out.csv"    # The folder and name of the CSV for your output.

Import-Module ReadMsgFile

Set-Location $msgRootFolder

$msgFiles = (Get-ChildItem -Filter "*.msg")

foreach ($msgFile in $msgFiles) {
    try {
        Read-MsgFile $msgFile | `
            Select-Object From, To, CC, Sent, Attachments, Subject, @{Name = 'FileName'; Expression = { $msgFile.Name } } | `
            Export-Csv $outputFile -Append -NoTypeInformation
    }
    catch {
    }
}