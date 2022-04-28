// Created by Elders in project VSE-FormatDocumentOnSave. 
// See https://github.com/Elders/VSE-FormatDocumentOnSave.

using System;
using System.Linq;

namespace Tinyfish.FormatOnSave
{
    public class AllowDenyDocumentFilter
    {
        private readonly Func<string, bool> _isAllowed = fileName => true;

        public AllowDenyDocumentFilter(string[] allowedExtensions, string[] deniedExtensions)
        {
            allowedExtensions = allowedExtensions?.Where(x => !x.Equals(".*") && !string.IsNullOrEmpty(x)).ToArray();
            deniedExtensions = deniedExtensions?.Where(x => !x.Equals(".*") && !string.IsNullOrEmpty(x)).ToArray();

            if (allowedExtensions != null && allowedExtensions.Any())
                _isAllowed = fileName => allowedExtensions.Any(ext => fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
            else if (deniedExtensions != null && deniedExtensions.Any())
                _isAllowed = fileName => !deniedExtensions.Any(ext => fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
        }

        public bool IsAllowed(string fileName)
        {
            return _isAllowed(fileName);
        }
    }
}
