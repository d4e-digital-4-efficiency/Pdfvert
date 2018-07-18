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
    internal static class OpenOfficeUtility
    {
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

        public static XComponent GetComponent()
        {
            return (XComponent)null;
        }

        public static XComponentContext GetComponentContext()
        {
            return Bootstrap.bootstrap();
        }

        public static XMultiServiceFactory GetMultiServiceFactory(XComponentContext componentContext)
        {
            return (XMultiServiceFactory)componentContext.getServiceManager();
        }

        public static XComponentLoader GetComponentLoader(XMultiServiceFactory multiServiceFactory)
        {
            return (XComponentLoader)multiServiceFactory.createInstance("com.sun.star.frame.Desktop");
        }

        public static XComponent InitDocument(XComponentLoader aLoader, string file, string target)
        {
            var openProps = new PropertyValue[] { new PropertyValue { Name = "Hidden", Value = new Any(true) } };
            var xComponent = aLoader.loadComponentFromURL(file, target, 0, openProps);

            return xComponent;
        }

        public static void SaveDocument(XComponent xComponent, string sourceFile, string destinationFile)
        {
            var filterType = sourceFile.ConvertExtensionToFilterType();
            var propertyValues = new PropertyValue[] {
                new PropertyValue { Name = "Overwrite", Value = new Any(true) }, // Setting the flag for overwriting
                new PropertyValue { Name = "FilterName", Value = new Any(filterType) }  // Setting the filter name
            };

            ((XStorable)xComponent).storeToURL(destinationFile, propertyValues);
        }
    }
}
