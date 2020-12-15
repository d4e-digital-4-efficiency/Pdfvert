using System;
using System.Diagnostics;
using System.Linq;
using uno;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;

namespace Pdfvert.Core.Utilities
{
    /// <summary>
    /// </summary>
    internal static class OpenOfficeUtility
    {
        /// <summary>
        /// Gets the component.
        /// </summary>
        /// <returns> The Component </returns>
        public static XComponent GetComponent()
        {
            return (XComponent)null;
        }

        /// <summary>
        /// Gets the component context.
        /// </summary>
        /// <returns> The component context </returns>
        public static XComponentContext GetComponentContext()
        {
            return Bootstrap.bootstrap();
        }

        /// <summary>
        /// Gets the component loader.
        /// </summary>
        /// <param name="multiServiceFactory"> The multi service factory. </param>
        /// <returns> The component loader </returns>
        public static XComponentLoader GetComponentLoader(XMultiServiceFactory multiServiceFactory)
        {
            return (XComponentLoader)multiServiceFactory.createInstance("com.sun.star.frame.Desktop");
        }

        /// <summary>
        /// Gets the multi service factory.
        /// </summary>
        /// <param name="componentContext"> The component context. </param>
        /// <returns> The multi service factory </returns>
        public static XMultiServiceFactory GetMultiServiceFactory(XComponentContext componentContext)
        {
            return (XMultiServiceFactory)componentContext.getServiceManager();
        }

        /// <summary>
        /// Initializes the document.
        /// </summary>
        /// <param name="aLoader"> a loader. </param>
        /// <param name="file"> The file. </param>
        /// <param name="target"> The target. </param>
        /// <returns> The Component </returns>
        public static XComponent InitDocument(XComponentLoader aLoader, string file, string target)
        {
            var openProps = new PropertyValue[] { new PropertyValue { Name = "Hidden", Value = new Any(true) } };
            var xComponent = aLoader.loadComponentFromURL(file, target, 0, openProps);

            return xComponent;
        }

        /// <summary>
        /// Saves the document.
        /// </summary>
        /// <param name="xComponent"> The x component. </param>
        /// <param name="sourceFile"> The source file. </param>
        /// <param name="destinationFile"> The destination file. </param>
        public static void SaveDocument(XComponent xComponent, string sourceFile, string destinationFile)
        {
            var filterType = sourceFile.ConvertExtensionToFilterType();
            var propertyValues = new PropertyValue[] {
                new PropertyValue { Name = "Overwrite", Value = new Any(true) }, // Setting the flag for overwriting
                new PropertyValue { Name = "FilterName", Value = new Any(filterType) }  // Setting the filter name
            };

            ((XStorable)xComponent).storeToURL(destinationFile, propertyValues);
        }

        /// <summary>
        /// Starts the open office.
        /// </summary>
        /// <exception cref="System.InvalidProgramException"> OpenOffice failed to start. </exception>
        public static void StartOpenOffice()
        {
            var ps = Process.GetProcessesByName("soffice");
            if (ps != null && ps.Any()) return;

            var pInfo = new ProcessStartInfo
            {
                FileName = "soffice.exe",
                Arguments = "-headless -nofirststartwizard",
                CreateNoWindow = true
            };

            var p = new Process { StartInfo = pInfo };
            var result = p.Start();

            if (result == false)
                throw new InvalidProgramException("OpenOffice failed to start.");
        }
    }
}