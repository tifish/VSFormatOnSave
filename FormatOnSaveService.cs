using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Tinyfish.FormatOnSave
{
    class FormatOnSaveService : IVsRunningDocTableEvents3
    {
        readonly DTE _dte;
        readonly IVsTextManager _textManager;
        readonly OptionsPage _optionsPage;
        readonly RunningDocumentTable _runningDocumentTable;

        public FormatOnSaveService(DTE dte, RunningDocumentTable runningDocumentTable, IVsTextManager textManager, OptionsPage optionsPage)
        {
            _runningDocumentTable = runningDocumentTable;
            _textManager = textManager;
            _dte = dte;
            _optionsPage = optionsPage;
        }

        public int OnBeforeSave(uint docCookie)
        {
            var document = FindDocument(docCookie);

            if (document == null)
            {
                return VSConstants.S_OK;
            }

            var isFilterAllowed = _optionsPage.AllowDenyFilter.IsAllowed(document);

            // Do TabToSpace before FormatDocument, since VS format may break the tab formatting.
            if (_optionsPage.EnableTabToSpace && isFilterAllowed)
            {
                TabToSpace(document);
            }
            if (_optionsPage.EnableRemoveAndSort && IsCsFile(document))
            {
                RemoveAndSort();
            }
            if (_optionsPage.EnableFormatDocument
                && _optionsPage.AllowDenyFormatDocumentFilter.IsAllowed(document))
            {
                FormatDocument();
            }
            // Do TabToSpace again after FormatDocument, since VS2017 may stick to tab. Should remove this after VS2017 fix the bug.
            if (_optionsPage.EnableTabToSpace && isFilterAllowed && _dte.Version == "15.0")
            {
                TabToSpace(document);
            }
            if (_optionsPage.EnableUnifyLineBreak && isFilterAllowed)
            {
                UnifyLineBreak(document);
            }
            if (_optionsPage.EnableUnifyEndOfFile && isFilterAllowed)
            {
                UnifyEndOfFile(document);
            }

            return VSConstants.S_OK;
        }

        static bool IsCsFile(Document document)
        {
            return document.FullName.EndsWith(".cs", StringComparison.OrdinalIgnoreCase);
        }

        void RemoveAndSort()
        {
            try
            {
                _dte.ExecuteCommand("Edit.RemoveAndSort", "");
            }
            catch (COMException)
            {

            }
        }

        void FormatDocument()
        {
            try
            {
                _dte.ExecuteCommand("Edit.FormatDocument", "");
            }
            catch (COMException)
            {

            }
        }

        void UnifyLineBreak(Document document)
        {
            var wpfTextView = GetWpfTextView(GetIVsTextView(document.FullName));

            var snapshot = wpfTextView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                var defaultLineBreak = "";
                switch (_optionsPage.LineBreak)
                {
                    case OptionsPage.LineBreakStyle.Unix:
                        defaultLineBreak = "\n";
                        break;
                    case OptionsPage.LineBreakStyle.Windows:
                        defaultLineBreak = "\r\n";
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                foreach (var line in snapshot.Lines)
                {
                    if (line.GetLineBreakText() == defaultLineBreak)
                    {
                        continue;
                    }

                    edit.Delete(line.End.Position, line.LineBreakLength);
                    edit.Insert(line.End.Position, defaultLineBreak);
                }

                edit.Apply();
            }
        }

        static void UnifyEndOfFile(Document document)
        {
            var wpfTextView = GetWpfTextView(GetIVsTextView(document.FullName));

            var snapshot = wpfTextView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                var lineNumber = snapshot.LineCount - 1;
                while (lineNumber >= 0 && snapshot.GetLineFromLineNumber(lineNumber).GetText().Trim() == "")
                {
                    lineNumber--;
                }

                var hasModified = false;
                var startEmptyLineNumber = lineNumber + 1;

                // Supply one empty line
                if (startEmptyLineNumber > snapshot.LineCount - 1)
                {
                    var firstLine = snapshot.GetLineFromLineNumber(1);
                    var defaultLineBreakText = firstLine.GetLineBreakText();
                    if (!string.IsNullOrEmpty(defaultLineBreakText))
                    {
                        var lastLine = snapshot.GetLineFromLineNumber(startEmptyLineNumber - 1);
                        edit.Insert(lastLine.End, defaultLineBreakText);
                        hasModified = true;
                    }
                }
                // Nothing to format
                else if (startEmptyLineNumber == snapshot.LineCount - 1)
                {
                    // do nothing
                }
                // Delete redudent empty lines
                else if (startEmptyLineNumber <= snapshot.LineCount - 1)
                {
                    var startPosition = snapshot.GetLineFromLineNumber(startEmptyLineNumber).Start.Position;
                    var endPosition = snapshot.GetLineFromLineNumber(snapshot.LineCount - 1).EndIncludingLineBreak.Position;
                    edit.Delete(startPosition, endPosition - startPosition);
                    hasModified = true;
                }

                if (hasModified)
                {
                    edit.Apply();
                }
            }
        }

        class SpaceStringPool
        {
            readonly string[] _stringCache = new string[8];

            public string GetString(int spaceCount)
            {
                if (spaceCount <= 0)
                {
                    throw new ArgumentOutOfRangeException();
                }

                var index = spaceCount - 1;

                if (spaceCount > _stringCache.Length)
                {
                    return new string(' ', spaceCount);
                }
                else if (_stringCache[index] == null)
                {
                    _stringCache[index] = new string(' ', spaceCount);
                    return _stringCache[index];
                }
                else
                {
                    return _stringCache[index];
                }
            }
        }

        readonly SpaceStringPool _spaceStringPool = new SpaceStringPool();

        void TabToSpace(Document document)
        {
            var wpfTextView = GetWpfTextView(GetIVsTextView(document.FullName));

            var snapshot = wpfTextView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                var hasModifed = false;

                foreach (var line in snapshot.Lines)
                {
                    var lineText = line.GetText();

                    if (!lineText.Contains('\t'))
                    {
                        continue;
                    }

                    var positionOffset = 0;

                    for (int i = 0; i < lineText.Length; i++)
                    {
                        var currentChar = lineText[i];
                        if (currentChar == '\t')
                        {
                            var absTabPosition = line.Start.Position + i;
                            edit.Delete(absTabPosition, 1);
                            var spaceCount = _optionsPage.TabToSpaceSize -
                                             (i + positionOffset) % _optionsPage.TabToSpaceSize;
                            edit.Insert(absTabPosition, _spaceStringPool.GetString(spaceCount));
                            positionOffset += spaceCount - 1;
                            hasModifed = true;
                        }
                        else if (IsCjkCharacter(currentChar))
                        {
                            positionOffset++;
                        }
                    }
                }

                if (hasModifed)
                {
                    edit.Apply();
                }
            }
        }

        readonly Regex _cjkRegex = new Regex(
            @"\p{IsHangulJamo}|" +
            @"\p{IsCJKRadicalsSupplement}|" +
            @"\p{IsCJKSymbolsandPunctuation}|" +
            @"\p{IsEnclosedCJKLettersandMonths}|" +
            @"\p{IsCJKCompatibility}|" +
            @"\p{IsCJKUnifiedIdeographsExtensionA}|" +
            @"\p{IsCJKUnifiedIdeographs}|" +
            @"\p{IsHangulSyllables}|" +
            @"\p{IsCJKCompatibilityForms}|" +
            @"\p{IsHalfwidthandFullwidthForms}");

        bool IsCjkCharacter(char character)
        {
            return _cjkRegex.IsMatch(character.ToString());
        }

        Document FindDocument(uint docCookie)
        {
            var documentInfo = _runningDocumentTable.GetDocumentInfo(docCookie);
            var documentPath = documentInfo.Moniker;

            return _dte.Documents.Cast<Document>().FirstOrDefault(doc => doc.FullName == documentPath);
        }

        static IWpfTextView GetWpfTextView(IVsTextView vTextView)
        {
            IWpfTextView view = null;
            var userData = vTextView as IVsUserData;

            if (null != userData)
            {
                var guidViewHost = DefGuidList.guidIWpfTextViewHost;
                userData.GetData(ref guidViewHost, out var holder);
                var viewHost = (IWpfTextViewHost)holder;
                view = viewHost.TextView;
            }

            return view;
        }

        static IVsTextView GetIVsTextView(string filePath)
        {
            var dte2 = (EnvDTE80.DTE2)Package.GetGlobalService(typeof(SDTE));
            var sp = (Microsoft.VisualStudio.OLE.Interop.IServiceProvider)dte2;
            var serviceProvider = new ServiceProvider(sp);
            return VsShellUtilities.IsDocumentOpen(serviceProvider, filePath, Guid.Empty, out var uiHierarchy, out uint itemId, out var windowFrame)
                ? VsShellUtilities.GetTextView(windowFrame) : null;
        }

        IVsOutputWindowPane _outputWindowPane;
        Guid _outputWindowPaneGuid = new Guid("8AEEC946-659A-4D14-8340-730B2A0FF39C");

        void OutputString(string message)
        {
            if (_outputWindowPane == null)
            {
                var outWindow = (IVsOutputWindow)Package.GetGlobalService(typeof(SVsOutputWindow));
                outWindow.CreatePane(ref _outputWindowPaneGuid, "VSFormatOnSave", 1, 1);
                outWindow.GetPane(ref _outputWindowPaneGuid, out _outputWindowPane);
            }

            _outputWindowPane.OutputString(message + Environment.NewLine);
            _outputWindowPane.Activate(); // Brings this pane into view
        }

        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRdtLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRdtLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterSave(uint docCookie)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }

        int IVsRunningDocTableEvents3.OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld,
            string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew)
        {
            return VSConstants.S_OK;
        }


        int IVsRunningDocTableEvents2.OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld,
            string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew)
        {
            return VSConstants.S_OK;
        }
    }
}
