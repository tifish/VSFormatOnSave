// Created by Elders in project VSE-FormatDocumentOnSave. 
// See https://github.com/Elders/VSE-FormatDocumentOnSave.

using System;
using System.IO;
using System.Linq;

namespace Tinyfish.FormatOnSave
{
    public class AllowDenyPathFilter
    {
        private string[] _whitelistedPaths { get; set; }
        private string[] _blacklistedPaths { get; set; }
        public AllowDenyPathFilter(string[] whitelistPaths, string[] blacklistPaths)
        {
            _whitelistedPaths = whitelistPaths.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
            _blacklistedPaths = blacklistPaths.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
        }

        // allowed if a path is whitelisted or not blacklisted
        public bool IsAllowed(string path)
        {
            if (string.IsNullOrEmpty(path))
                return true;

            bool whiteListed = _whitelistedPaths.Any(x => IsUnderFolder(x, path));
            if (whiteListed)
            {
                return true;
            }
            else
            {
                return !_blacklistedPaths.Any(x => IsUnderFolder(x, path));
            }
        }

        private bool CheckIfParentDir(DirectoryInfo childDir, DirectoryInfo parentDir)
        {
            if (!parentDir.Exists)
                return false;

            bool isParent = false;
            while (childDir.Parent != null)
            {
                if (childDir.Parent.FullName == parentDir.FullName)
                {
                    isParent = true;
                    break;
                }
                else
                {
                    childDir = childDir.Parent;
                }
            }
            return isParent;
        }

        private bool IsUnderFolder(string parent, string child)
        {
            try
            {
                DirectoryInfo parentDir = new DirectoryInfo(parent.TrimEnd(Path.DirectorySeparatorChar));
                DirectoryInfo childDir = new DirectoryInfo(child.TrimEnd(Path.DirectorySeparatorChar));

                if (parentDir.FullName == childDir.FullName)
                {
                    return true;
                }

                return CheckIfParentDir(childDir, parentDir);
            }
            catch
            {
                return false;
            }
        }
    }

    public class AllowDenyDocumentFilter
    {
        private readonly Func<string, string, AllowDenyPathFilter, bool> _isAllowed = (fileName, path, pathFilter) => pathFilter.IsAllowed(path);
        private AllowDenyPathFilter _pathFilter;

        public AllowDenyDocumentFilter(string[] allowedExtensions, string[] deniedExtensions, AllowDenyPathFilter pathfilter)
        {
            allowedExtensions = allowedExtensions?.Where(x => !x.Equals(".*") && !string.IsNullOrEmpty(x)).ToArray();
            deniedExtensions = deniedExtensions?.Where(x => !x.Equals(".*") && !string.IsNullOrEmpty(x)).ToArray();
            _pathFilter = pathfilter;

            if (allowedExtensions != null && allowedExtensions.Any())
                _isAllowed = (fileName, path, _pathFilter) => allowedExtensions.Any(ext => fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) && _pathFilter.IsAllowed(path);
            else if (deniedExtensions != null && deniedExtensions.Any())
                _isAllowed = (fileName, path, _pathFilter) => !deniedExtensions.Any(ext => fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) && _pathFilter.IsAllowed(path);
        }

        public bool IsAllowed(string fileName, string path)
        {
            return _isAllowed(fileName, path, _pathFilter);
        }
    }
}
