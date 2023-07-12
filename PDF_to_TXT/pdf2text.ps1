#Unzip pdf2text.zip to your drive. Somewhere like C:\pdf2text 

$exeLocation = "C:\pdf2text"            # The folder where you put the EXE.
$exeName = "pdf2text.exe"               # Leave this alone unless you rename the EXE.
$pdfRootFolder = "C:\PDFs"              # The top level folder of the file structure where the PDFs are.
$outputFolder = "C:\My\Folder\"         # The folder for your text file output.

$exePath = "$exeLocation"+"\"+"$exeName"

& "$exePath" --file $pdfRootFolder --subfolders --output $outputFolder