using NLog;
using System;

namespace Pdfvert.Core.Utilities
{
    /// <summary>
    /// Log Utility
    /// </summary>
    internal static class LogUtility
    {
        /// <summary>
        /// Logger of Logger
        /// </summary>
        private static Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Log the the specified message errors .
        /// </summary>
        /// <param name="message"> The message. </param>
        public static void Error(string message)
        {
            Console.WriteLine(message);
            logger.Error(message);
        }

        /// <summary>
        /// Logs the specified message.
        /// </summary>
        /// <param name="message"> The message. </param>
        public static void Log(string message)
        {
            Console.WriteLine(message);
            logger.Info(message);
        }
    }
}