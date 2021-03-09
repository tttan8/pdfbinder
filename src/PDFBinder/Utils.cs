using System;
using System.Collections.Generic;
using System.Text;

namespace PDFBinder
{
    public static class Utils
    {
        public static string GetFileNameFromPath(string path)
        {
            var startIndex = path.LastIndexOf("\\");
            var endIndex = path.LastIndexOf(".");
            return path.Substring(startIndex + 1, endIndex - startIndex - 1);
        }
    }
}
