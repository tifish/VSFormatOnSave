// Created by Elders in project VSE-FormatDocumentOnSave. 
// See https://github.com/Elders/VSE-FormatDocumentOnSave.

using EnvDTE;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Tinyfish.FormatOnSave
{
    public class AllowDenyDocumentFilter
    {
        readonly Func<Document, bool> _isAllowed = doc => true;

        /// <summary>
        /// Everything is allowed when this ctor is used.
        /// </summary>
        public AllowDenyDocumentFilter() { }

        public AllowDenyDocumentFilter(IEnumerable<string> allowedExtensions, IEnumerable<string> deniedExtensions)
        {
            allowedExtensions = allowedExtensions.Where(x => x.Equals(".*") == false && string.IsNullOrEmpty(x) == false);
            deniedExtensions = deniedExtensions.Where(x => x.Equals(".*") == false && string.IsNullOrEmpty(x) == false);

            if (allowedExtensions.Count() > 0)
            {
                _isAllowed = doc => allowedExtensions.Any(ext => doc.FullName.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
            }
            else if (deniedExtensions.Count() > 0)
            {
                _isAllowed = doc => deniedExtensions.Any(ext => doc.FullName.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) == false;
            }
        }

        public bool IsAllowed(Document document)
        {
            return _isAllowed(document);
        }
    }
}
