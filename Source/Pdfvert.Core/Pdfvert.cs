using Pdfvert.Core.Utilities;
using System;
using System.Threading;

namespace Pdfvert.Core
{
    /// <summary>
    /// Pdfvert Service for convert Office document to PDF
    /// </summary>
    public static class PdfvertService
    {
        /// <summary>
        /// Converts the file to PDF.
        /// </summary>
        /// <param name="inputFile"> The input file. </param>
        /// <param name="outputFile"> The output file. </param>
        /// <exception cref="System.InvalidProgramException">
        /// These aren't the file types you're looking for. Invalid file type for source file. or
        /// These aren't the file types you're looking for. Invalid file type for source file. Must
        /// be PDF.
        /// </exception>
        public static void ConvertFileToPdf(string inputFile, string outputFile)
        {
            PrintStart(inputFile, outputFile);

            if (!inputFile.IsValidExtension())
                throw new InvalidProgramException($"These aren't the file types you're looking for. Invalid file type for source file.");

            if (!outputFile.IsPdf())
                throw new InvalidProgramException($"These aren't the file types you're looking for. Invalid file type for source file. Must be PDF.");

            var xComponent = OpenOfficeUtility.GetComponent();

            try
            {
                LogUtility.Log("Starting Open Office...");
                OpenOfficeUtility.StartOpenOffice();

                LogUtility.Log("...Getting a ComponentContext");
                var xLocalContext = OpenOfficeUtility.GetComponentContext();

                LogUtility.Log("...Getting a MultiServiceFactory");
                var xRemoteFactory = OpenOfficeUtility.GetMultiServiceFactory(xLocalContext);

                LogUtility.Log("...Getting a ComponentLoader");
                var aLoader = OpenOfficeUtility.GetComponentLoader(xRemoteFactory);

                LogUtility.Log("...Loading source file");
                xComponent = OpenOfficeUtility.InitDocument(aLoader, inputFile.ConvertPath(), "_blank");

                LogUtility.Log("...wait for loading");
                while (xComponent == null)
                    Thread.Sleep(1000);

                LogUtility.Log("...saving/exporting the document");
                OpenOfficeUtility.SaveDocument(xComponent, inputFile, outputFile.ConvertPath());
            }
            finally
            {
                LogUtility.Log("...disposing ComponentLoader");
                PrintEnd();
                if (xComponent != null) xComponent.dispose();
            }
        }

        /// <summary>
        /// Prints the end log.
        /// </summary>
        private static void PrintEnd()
        {
            LogUtility.Log("");
            LogUtility.Log("You're PDF Is ready!");
            LogUtility.Log("***************************");
            LogUtility.Log("");
        }

        /// <summary>
        /// Prints the start log.
        /// </summary>
        /// <param name="inputFile"> The input file. </param>
        /// <param name="outputFile"> The output file. </param>
        private static void PrintStart(string inputFile, string outputFile)
        {
            LogUtility.Log("***************************");
            LogUtility.Log("    Starting Conversion ");
            LogUtility.Log("***************************");
            LogUtility.Log(DateTime.Now.ToString());
            LogUtility.Log($"Input File: {inputFile}");
            LogUtility.Log($"Ouput FIle: {outputFile}");
        }
    }
}