# Pdfvert
---
**Convert a variety of file types to PDF using OpenOffice v4.1.8. (released on November 10, 2020)**

Uses OpenOffice SDK v4.1.8 (released on November 10, 2020) - DLLs included in NuGet Package (no need to download)

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

#### Download OpenOffice here:
[http://www.openoffice.org/download/](http://www.openoffice.org/download/)

#### Download OpenOffice SDK here:
[https://www.openoffice.org/download/devbuilds.html](https://www.openoffice.org/download/devbuilds.html)

#### Install nuget package:
[https://www.nuget.org/packages/Pdfvert/](https://www.nuget.org/packages/Pdfvert/)
```powershell
PM> Install-Package Pdfvert
```

#### Usage:
```C#
var inputFile = "C:\temp\test.docx"; //file you want to convert
var outputFile = "C:\temp\test.pdf"; //new path and new file name for converted PDF (must have .pdf extension)

PdfvertService.ConvertFileToPdf(inputFile, outputFile);
```
