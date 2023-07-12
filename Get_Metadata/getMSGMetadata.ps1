#This will require downloading and installing a PS module from the gallery. It will prompt you to do so. 

$msgRootFolder = "C:\temp"              # The top level folder of the file structure where the MSG files are.
$outputFile = "C:\My\Folder\out.csv"    # The folder and name of the CSV for your output.
Import-Module ReadMsgFile               # This will download the needed module

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