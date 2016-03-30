using System;
using System.Threading;
using Pdfvert.Utilities;

namespace Pdfvert
{
    public static class Pdfvert
    {
        public static void ConvertFile(string inputFile, string outputFile)
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

        private static void PrintStart(string inputFile, string outputFile)
        {
            LogUtility.Log("***************************");
            LogUtility.Log("    Starting Conversion ");
            LogUtility.Log("***************************");
            LogUtility.Log(DateTime.Now.ToString());
            LogUtility.Log($"Input File: {inputFile}");
            LogUtility.Log($"Ouput FIle: {outputFile}");
        }

        private static void PrintEnd()
        {
            LogUtility.Log("");
            LogUtility.Log("Daaaaaamn Daniel - You're PDF Is ready!");
            LogUtility.Log("***************************");
            LogUtility.Log("");
        }
    }
}
