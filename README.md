# Pdfvert
Convert a variety of file types to PDF using OpenOffice

#### Supported File Types:
* .doc
* .docx
* .txt
* .rtf
* .html
* .htm
* .xml
* .odt
* .wps
* .wpd
* .xls
* .xlsb
* .xlsx
* .ods
* .ppt
* .pptx
* .odp
* 
#### Download OpenOffice here:
[http://www.openoffice.org/download/](http://www.openoffice.org/download/)

#### Install nuget package:
```powershell
PM> Install-Package Pdfver
```

#### Usage:
```C#
var inputFile = ""; //file you want to convert
var outputFile = ""; //new path and new file name for converted PDF (must have .pdf extension)

PdfvertService.ConvertFileToPdf(inputFile, outputFile);
```

