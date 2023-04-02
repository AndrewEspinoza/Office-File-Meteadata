$exeLocation = "C:\My\Folder"           # The folder where you put the EXE.
$exeName = "exiftool.exe"               # Leave this alone unless you rename the EXE.
$pdfRootFolder = "C:\PDFs"              # The top level folder of the file structure where the PDFs are.
$outputFile = "C:\My\Folder\out.csv"    # The folder and name of the CSV for your output.

$exePath = "$exeLocation"+"\"+"$exeName"

& "$exePath" $pdfRootFolder -ext pdf -r -csv > $outputFile