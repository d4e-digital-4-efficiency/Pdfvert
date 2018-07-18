using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Pdfvert.Core.Utilities;

namespace Pdfvert.Tests
{
    [TestClass]
    public class OpenOfficeUtilityTests
    {
        [TestMethod]
        public void TestMethod1()
        {
            var startingProcesses = Process.GetProcessesByName("soffice");
            var startingCount = startingProcesses.Length;

            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();
            OpenOfficeUtility.StartOpenOffice();

            var endingProcesses = Process.GetProcessesByName("soffice");
            var endingCount = endingProcesses.Length;

            Console.WriteLine($"Starting count: {startingCount}");
            Console.WriteLine($"Ending count: {endingCount}");

            Assert.IsTrue(endingCount >= startingCount);
            Assert.IsTrue((endingCount == startingCount) || (endingCount == startingCount + 1));
        }
    }
}
