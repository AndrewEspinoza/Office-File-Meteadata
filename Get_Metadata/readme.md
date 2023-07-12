# Pulling Metadata from MS Office files and PDFs

This process uses [ExifTool by Phil Harvey](https://exiftool.org/).

## Usage

1. Download the EXE and place on your local machine. This is a standalone executable and does not need to be installed.
1. Update the scripts found in here with your local file locations and run it. 
1. The output will not be a natively readable CSV file. You can import the file into Excel and have it parse the file to make it usable. 


# Pulling Metadata from Email MSG files
In order to pull metadata from MSG files using PowerShell, you need to install a module that can read MSG files. 

I was able to find __ReadMsgFile__ in the PowerShell gallery for this. https://www.powershellgallery.com/packages/ReadMsgFile/2.1

Install using `Install-Module -Name ReadMsgFile`