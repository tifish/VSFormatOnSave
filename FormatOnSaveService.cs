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

        IWpfTextView GetWpfTextView(IVsTextView vTextView)
        {
            IWpfTextView view = null;
            var userData = (IVsUserData)vTextView;

            if (null != userData)
            {
                var guidViewHost = DefGuidList.guidIWpfTextViewHost;
                userData.GetData(ref guidViewHost, out var holder);
                var viewHost = (IWpfTextViewHost)holder;
                view = viewHost.TextView;
            }

            return view;
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
            if (_settingsPage.EnableUnifyLineBreak
                && _settingsPage.AllowDenyUnifyLineBreakFilter.IsAllowed(document))
            {
                UnifyLineBreak();
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

        void UnifyLineBreak()
        {
            _textManager.GetActiveView(1, null, out var vsTextView);
            var wpfTextView = GetWpfTextView(vsTextView);

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

        private Document FindDocument(uint docCookie)
        {
            var documentInfo = _runningDocumentTable.GetDocumentInfo(docCookie);
            var documentPath = documentInfo.Moniker;

            return _dte.Documents.Cast<Document>().FirstOrDefault(doc => doc.FullName == documentPath);
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








