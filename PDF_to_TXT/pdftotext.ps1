$exeLocation = "C:\pdftotext"            # The folder where you put the EXE.
$exeName = "pdftotext.exe"               # Leave this alone unless you rename the EXE.
$pdfRootFolder = "C:\PDFs"              # The top level folder of the file structure where the PDFs are.
$outputFolder = "C:\My\Folder\"         # The folder for your text file output.

$exePath = "$exeLocation"+"\"+"$exeName"

& "$exePath" $pdfRootFolder $outputFolder