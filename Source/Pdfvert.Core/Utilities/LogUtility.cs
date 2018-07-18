using System;
using NLog;

namespace Pdfvert.Core.Utilities
{
    internal static class LogUtility
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public static void Error(string message)
        {
            Console.WriteLine(message);
            logger.Error(message);
        }

        public static void Log(string message)
        {
            Console.WriteLine(message);
            logger.Info(message);
        }
    }
}

