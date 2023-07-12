$exeLocation = "C:\My\Folder"           # The folder where you put the EXE.
$exeName = "exiftool.exe"               # Leave this alone unless you rename the EXE.
$rootFolder = "C:\MyFiles"              # The top level folder of the file structure where the source files are.
$outputFile = "C:\My\Folder\out.csv"    # The folder and name of the CSV for your output.

$exePath = "$exeLocation"+"\"+"$exeName"

& "$exePath" $rootFolder -ext ppt -ext+ pptx -ext+ pptm  -r -csv > $outputFile