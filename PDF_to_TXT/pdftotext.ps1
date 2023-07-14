$exeLocation = "C:\temp"           # The folder where you put the EXE.
$exeName = "pdftotext.exe"         # Leave this alone unless you rename the EXE.
$pdfRootFolder = "C:\temp\*"       # MUST INCLUDE WILDCARD "*". The top level folder of the file structure where the PDFs are.
$outputFolder = "C:\temp2"         # The folder for your text file output.



$pdfs = (Get-ChildItem -Path $pdfRootFolder -Include *.pdf -Recurse) # Gets all the PDFs. 
$exePath = "$exeLocation"+"\"+"$exeName"                             # Execution path.

foreach ($pdf in $pdfs) {
     $pdfPath = $pdf.Fullname                              # Gets the full file path and the file name.
     $txtPath = "$outputFolder"+"\"+$pdf.Name+".txt"       # Defines the full file path and file name of the text file output.
     & "$exePath" $pdfPath $txtPath                        # Converts the PDf to TXT.
     }