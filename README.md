# Pdfvert
Convert a variety of file types to PDF using OpenOffice

Download OpenOffice here:
http://www.openoffice.org/download/

Install nuget package:
PM> Install-Package Pdfver

Usage:

var inputFile = ""; //file you want to convert
var outputFile = ""; //new path and new file name for converted PDF (must have .pdf extension)

PdfvertService.ConvertFileToPdf(inputFile, outputFile);

