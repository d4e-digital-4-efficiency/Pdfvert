using System;
using System.IO;

namespace Pdfvert.Core
{
    internal static class Extensions
    {
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

        public static bool IsValidExtension(this string path)
        {
            return !string.IsNullOrEmpty(path.ConvertExtensionToFilterType());
        }

        public static bool IsPdf(this string path)
        {
            var extension = Path.GetExtension(path);
            return !string.IsNullOrEmpty(extension) && extension.ToLower() == ".pdf";
        }

        public static string ConvertPath(this string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new NullReferenceException("Null or empty path passed to OpenOffice");

            return string.Format("file:///{0}", path.Replace(@"\", "/"));
        }
    }
}
