using System;
using System.IO;

namespace Pdfvert.Core
{
    /// <summary>
    /// Extensions
    /// </summary>
    internal static class Extensions
    {
        /// <summary>
        /// Converts the type of the extension to filter.
        /// </summary>
        /// <param name="path"> The path. </param>
        /// <returns> String extension </returns>
        public static string ConvertExtensionToFilterType(this string path)
        {
            switch (Path.GetExtension(path).ToLower())
            {
                case ".doc":
                case ".docx":
                case ".txt":
                case ".rtf":
                case ".html":
                case ".htm":
                case ".xml":
                case ".odt":
                case ".wps":
                case ".wpd":
                    return "writer_pdf_Export";

                case ".xls":
                case ".xlsb":
                case ".xlsx":
                case ".ods":
                    return "calc_pdf_Export";

                case ".ppt":
                case ".pptx":
                case ".odp":
                    return "impress_pdf_Export";

                default:
                    return null;
            }
        }

        /// <summary>
        /// Converts the path.
        /// </summary>
        /// <param name="path"> The path. </param>
        /// <returns> The file path </returns>
        /// <exception cref="System.NullReferenceException">
        /// Null or empty path passed to OpenOffice
        /// </exception>
        public static string ConvertPath(this string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new NullReferenceException("Null or empty path passed to OpenOffice");

            return string.Format("file:///{0}", path.Replace(@"\", "/"));
        }

        /// <summary>
        /// Determines whether this instance is PDF.
        /// </summary>
        /// <param name="path"> The path. </param>
        /// <returns> <c> true </c> if the specified path is PDF; otherwise, <c> false </c>. </returns>
        public static bool IsPdf(this string path)
        {
            var extension = Path.GetExtension(path);
            return !string.IsNullOrEmpty(extension) && extension.ToLower() == ".pdf";
        }

        /// <summary>
        /// Determines whether [is valid extension].
        /// </summary>
        /// <param name="path"> The path. </param>
        /// <returns>
        /// <c> true </c> if [is valid extension] [the specified path]; otherwise, <c> false </c>.
        /// </returns>
        public static bool IsValidExtension(this string path)
        {
            return !string.IsNullOrEmpty(path.ConvertExtensionToFilterType());
        }
    }
}