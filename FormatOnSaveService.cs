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
using Microsoft.VisualStudio.Text;

namespace Tinyfish.FormatOnSave
{
    class FormatOnSaveService : IVsRunningDocTableEvents3
    {
        readonly DTE _dte;
        readonly IVsTextManager _textManager;
        readonly SettingsPage _settingsPage;
        readonly RunningDocumentTable _runningDocumentTable;

        public FormatOnSaveService(DTE dte, RunningDocumentTable runningDocumentTable, IVsTextManager textManager, SettingsPage settingsPage)
        {
            _runningDocumentTable = runningDocumentTable;
            _textManager = textManager;
            _dte = dte;
            _settingsPage = settingsPage;
        }

        public int OnBeforeSave(uint docCookie)
        {
            var document = FindDocument(docCookie);

            if (document == null)
            {
                return VSConstants.S_OK;
            }

            if (_settingsPage.EnableRemoveAndSort && IsCsFile(document))
            {
                RemoveAndSort();
            }
            if (_settingsPage.EnableFormatDocument
                && _settingsPage.AllowDenyFormatDocumentFilter.IsAllowed(document))
            {
                FormatDocument();
            }
            var isFilterAllowed = _settingsPage.AllowDenyFilter.IsAllowed(document);
            if (_settingsPage.EnableUnifyLineBreak && isFilterAllowed)
            {
                UnifyLineBreak(document);
            }
            if (_settingsPage.EnableUnifyEndOfFile && isFilterAllowed)
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
                switch (_settingsPage.LineBreak)
                {
                    case SettingsPage.LineBreakStyle.Unix:
                        defaultLineBreak = "\n";
                        break;
                    case SettingsPage.LineBreakStyle.Windows:
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
                    edit.Delete(new Span(startPosition, endPosition - startPosition));
                    hasModified = true;
                }

                if (hasModified)
                {
                    edit.Apply();
                }
            }
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
