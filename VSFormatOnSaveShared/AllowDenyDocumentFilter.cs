// Created by Elders in project VSE-FormatDocumentOnSave. 
// See https://github.com/Elders/VSE-FormatDocumentOnSave.

using System;
using System.Collections.Generic;
using System.Linq;

namespace Tinyfish.FormatOnSave
{
    public class AllowDenyDocumentFilter
    {
        readonly Func<string, bool> _isAllowed = fileName => true;

        /// <summary>
        /// Everything is allowed when this ctor is used.
        /// </summary>
        public AllowDenyDocumentFilter() { }

        public AllowDenyDocumentFilter(IEnumerable<string> allowedExtensions, IEnumerable<string> deniedExtensions)
        {
            allowedExtensions = allowedExtensions.Where(x => x.Equals(".*") == false && string.IsNullOrEmpty(x) == false);
            deniedExtensions = deniedExtensions.Where(x => x.Equals(".*") == false && string.IsNullOrEmpty(x) == false);

            if (allowedExtensions.Any())
            {
                _isAllowed = fileName => allowedExtensions.Any(ext => fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
            }
            else if (deniedExtensions.Any())
            {
                _isAllowed = fileName => deniedExtensions.Any(ext => fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) == false;
            }
        }

        public bool IsAllowed(string fileName)
        {
            return _isAllowed(fileName);
        }
    }
}
