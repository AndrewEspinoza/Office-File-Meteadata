<#
.Synopsis
   Pulls the meta data from an Excel workbook.

.DESCRIPTION
   Pulls the meta data from the workbook itself since the file properties may not be updated.

.NOTES
   I could not find a way to make Excel update the file properties so I had to find this way to do it. 
   This is originally taken from https://devblogs.microsoft.com/scripting/hey-scripting-guy-how-can-i-read-microsoft-excel-metadata/

   Created on:   3-30-23
   Created by:   Andrew Espinoza
   Last Updated: 3-30-23
   Filename:     getExcelData.ps1

#>

function Get-ExcelData

{
#Add your file path here
$path = "C:\temp\Excel\"
Get-ChildItem -Path $path -File -Recurse | Select-Object Name

$file = "C:\temp\Excel\Excel1.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($file)
$binding = "System.Reflection.BindingFlags" -as [type]
Foreach($property in $workbook.BuiltInDocumentProperties)
{
  $pn = [System.__ComObject].invokemember("name",$binding::GetProperty,$null,$property,$null)
  trap [system.exception]
   {
     #write-host -foreground blue "Value not found for $pn"
    continue
   }
  "$pn`: " + [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$property,$null)
 }
$excel.quit()
}