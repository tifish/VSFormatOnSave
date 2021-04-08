using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using DefGuidList = Microsoft.VisualStudio.Editor.DefGuidList;
using IServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;
using Task = System.Threading.Tasks.Task;

namespace Tinyfish.FormatOnSave
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideAutoLoad("{f1536ef8-92ec-443c-9ed7-fdadf150da82}", PackageAutoLoadFlags.BackgroundLoad)] //To set the UI context to autoload a VSPackage
    [Guid(GuidList.GuidFormatOnSavePkgString)]
    [ProvideOptionPage(typeof(OptionsPage), "Format on Save", "Settings", 0, 0, true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public class FormatOnSavePackage : AsyncPackage
    {
        public DTE2 Dte;
        public OptionsPage OptionsPage;
        RunningDocumentTable _runningDocumentTable;
        ServiceProvider _serviceProvider;
        ITextUndoHistoryRegistry _undoHistoryRegistry;
        public OleMenuCommandService MenuCommandService;
        SolutionExplorerContextMenu _solutionExplorerContextMenu;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _runningDocumentTable = new RunningDocumentTable(this);
            OptionsPage = (OptionsPage) GetDialogPage(typeof(OptionsPage));

            Dte = await GetServiceAsync(typeof(SDTE)) as DTE2;
            _serviceProvider = new ServiceProvider((IServiceProvider) Dte);
            var componentModel = (IComponentModel) GetGlobalService(typeof(SComponentModel));
            _undoHistoryRegistry = componentModel.DefaultExportProvider.GetExportedValue<ITextUndoHistoryRegistry>();
            MenuCommandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            _solutionExplorerContextMenu = new SolutionExplorerContextMenu(this);

            var plugin = new VsRunningDocTableEventsHandler(this);
            _runningDocumentTable.Advise(plugin);
        }

        public void Format(uint docCookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var document = FindDocument(docCookie);
            Format(document);
        }

        Document FindDocument(uint docCookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var documentInfo = _runningDocumentTable.GetDocumentInfo(docCookie);
            var documentPath = documentInfo.Moniker;

            foreach (Document doc in Dte.Documents)
            {
                if (doc.FullName == documentPath)
                    return doc;
            }

            return null;
        }

        public bool Format(Document document)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (document == null || document.Type != "Text" || document.Language == null ||
                document.Language == "Plain Text")
                return false;

            var oldActiveDocument = Dte.ActiveDocument;
            document.Activate();

            try
            {
                var languageOptions = Dte.Properties["TextEditor", document.Language];
                var insertTabs = (bool) languageOptions.Item("InsertTabs").Value;
                var isFilterAllowed = OptionsPage.AllowDenyFilter.IsAllowed(document.Name);

                var vsTextView = GetIVsTextView(document.FullName);
                if (vsTextView == null)
                    return false;
                var wpfTextView = GetWpfTextView(vsTextView);
                if (wpfTextView == null)
                    return false;

                _undoHistoryRegistry.TryGetHistory(wpfTextView.TextBuffer, out var history);

                using (var undo = history?.CreateTransaction("Format on save"))
                {
                    vsTextView.GetCaretPos(out var oldCaretLine, out var oldCaretColumn);
                    vsTextView.SetCaretPos(oldCaretLine, 0);

                    // Do TabToSpace before FormatDocument, since VS format may break the tab formatting.
                    if (OptionsPage.EnableTabToSpace && isFilterAllowed && !insertTabs)
                        TabToSpace(wpfTextView, document.TabSize);

                    if (OptionsPage.EnableRemoveAndSort && IsCsFile(document))
                    {
                        if (!OptionsPage.EnableSmartRemoveAndSort || !HasIfCompilerDirective(wpfTextView))
                            RemoveAndSort();
                    }

                    if (OptionsPage.EnableFormatDocument &&
                        OptionsPage.AllowDenyFormatDocumentFilter.IsAllowed(document.Name))
                        FormatDocument();

                    // Do TabToSpace again after FormatDocument, since VS2017 may stick to tab. Should remove this after VS2017 fix the bug.
                    if (OptionsPage.EnableTabToSpace && isFilterAllowed && !insertTabs && Dte.Version == "15.0" &&
                        document.Language == "C/C++")
                        TabToSpace(wpfTextView, document.TabSize);

                    if (OptionsPage.EnableUnifyLineBreak && isFilterAllowed)
                        UnifyLineBreak(wpfTextView);

                    if (OptionsPage.EnableUnifyEndOfFile && isFilterAllowed)
                        UnifyEndOfFile(wpfTextView);

                    if (OptionsPage.EnableForceUtf8WithBom &&
                        OptionsPage.AllowDenyForceUtf8WithBomFilter.IsAllowed(document.Name))
                        ForceUtf8WithBom(wpfTextView);

                    if (OptionsPage.EnableRemoveTrailingSpaces && isFilterAllowed &&
                        (Dte.Version == "11.0" || !OptionsPage.EnableFormatDocument))
                        RemoveTrailingSpaces(wpfTextView);

                    vsTextView.GetCaretPos(out var newCaretLine, out var newCaretColumn);
                    vsTextView.SetCaretPos(newCaretLine, oldCaretColumn);

                    undo?.Complete();
                }
            }
            finally
            {
                oldActiveDocument?.Activate();
            }

            return true;
        }

        static bool IsCsFile(Document document)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return document.FullName.EndsWith(".cs", StringComparison.OrdinalIgnoreCase);
        }

        static bool HasIfCompilerDirective(ITextView wpfTextView)
        {
            return wpfTextView.TextSnapshot.GetText().Contains("#if");
        }

        void RemoveAndSort()
        {
            try
            {
                Dte.ExecuteCommand("Edit.RemoveAndSort", string.Empty);
            }
            catch (COMException)
            {
            }
        }

        void FormatDocument()
        {
            try
            {
                Dte.ExecuteCommand("Edit.FormatDocument", string.Empty);
            }
            catch (COMException)
            {
            }
        }

        void UnifyLineBreak(ITextView wpfTextView)
        {
            var snapshot = wpfTextView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                string defaultLineBreak;
                switch (OptionsPage.LineBreak)
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
                    // if line break is defaultLineBreak or the line is the last => continue;
                    if (line.GetLineBreakText() == defaultLineBreak || line.LineNumber == snapshot.LineCount - 1)
                        continue;

                    edit.Delete(line.End.Position, line.LineBreakLength);
                    edit.Insert(line.End.Position, defaultLineBreak);
                }

                edit.Apply();
            }
        }

        static void UnifyEndOfFile(ITextView textView)
        {
            var snapshot = textView.TextSnapshot;
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
                    var endPosition = snapshot.GetLineFromLineNumber(snapshot.LineCount - 1).EndIncludingLineBreak
                        .Position;
                    edit.Delete(startPosition, endPosition - startPosition);
                    hasModified = true;
                }

                if (hasModified)
                    edit.Apply();
            }
        }

        static void ForceUtf8WithBom(ITextView wpfTextView)
        {
            try
            {
                ITextDocument textDocument;
                wpfTextView.TextDataModel.DocumentBuffer.Properties.TryGetProperty(typeof(ITextDocument),
                    out textDocument);

                if (textDocument.Encoding.EncodingName != Encoding.UTF8.EncodingName)
                    textDocument.Encoding = Encoding.UTF8;
            }
            catch (Exception)
            {
            }
        }

        static void RemoveTrailingSpaces(ITextView textView)
        {
            var snapshot = textView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                var hasModified = false;

                for (var i = 0; i < snapshot.LineCount; i++)
                {
                    var line = snapshot.GetLineFromLineNumber(i);
                    var lineText = line.GetText();

                    var trimmedLength = lineText.TrimEnd().Length;
                    if (trimmedLength == lineText.Length)
                        continue;

                    var spaceLength = lineText.Length - trimmedLength;
                    var endPosition = line.End.Position;
                    edit.Delete(endPosition - spaceLength, spaceLength);
                    hasModified = true;
                }

                if (hasModified)
                    edit.Apply();
            }
        }

        class SpaceStringPool
        {
            readonly string[] _stringCache = new string[8];

            public string GetString(int spaceCount)
            {
                if (spaceCount <= 0)
                    throw new ArgumentOutOfRangeException();

                var index = spaceCount - 1;

                if (spaceCount > _stringCache.Length)
                    return new string(' ', spaceCount);
                if (_stringCache[index] == null)
                {
                    _stringCache[index] = new string(' ', spaceCount);
                    return _stringCache[index];
                }

                return _stringCache[index];
            }
        }

        readonly SpaceStringPool _spaceStringPool = new SpaceStringPool();

        void TabToSpace(ITextView wpfTextView, int tabSize)
        {
            var snapshot = wpfTextView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                var hasModifed = false;

                foreach (var line in snapshot.Lines)
                {
                    var lineText = line.GetText();

                    if (!lineText.Contains('\t'))
                        continue;

                    var positionOffset = 0;

                    for (var i = 0; i < lineText.Length; i++)
                    {
                        var currentChar = lineText[i];
                        if (currentChar == '\t')
                        {
                            var absTabPosition = line.Start.Position + i;
                            edit.Delete(absTabPosition, 1);
                            var spaceCount = tabSize - (i + positionOffset) % tabSize;
                            edit.Insert(absTabPosition, _spaceStringPool.GetString(spaceCount));
                            positionOffset += spaceCount - 1;
                            hasModifed = true;
                        }
                        else if (IsCjkCharacter(currentChar))
                            positionOffset++;
                    }
                }

                if (hasModifed)
                    edit.Apply();
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

        static IWpfTextView GetWpfTextView(IVsTextView vTextView)
        {
            IWpfTextView view = null;
            var userData = (IVsUserData) vTextView;

            if (userData != null)
            {
                var guidViewHost = DefGuidList.guidIWpfTextViewHost;
                userData.GetData(ref guidViewHost, out var holder);
                var viewHost = (IWpfTextViewHost) holder;
                view = viewHost.TextView;
            }

            return view;
        }

        IVsTextView GetIVsTextView(string filePath)
        {
            return VsShellUtilities.IsDocumentOpen(_serviceProvider, filePath, Guid.Empty, out var uiHierarchy,
                out var itemId, out var windowFrame)
                ? VsShellUtilities.GetTextView(windowFrame)
                : null;
        }

        IVsOutputWindowPane _outputWindowPane;
        Guid _outputWindowPaneGuid = new Guid("8AEEC946-659A-4D14-8340-730B2A0FF39C");

        public void OutputString(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_outputWindowPane == null)
            {
                var outWindow = (IVsOutputWindow) GetGlobalService(typeof(SVsOutputWindow));
                outWindow.CreatePane(ref _outputWindowPaneGuid, "VSFormatOnSave", 1, 1);
                outWindow.GetPane(ref _outputWindowPaneGuid, out _outputWindowPane);
            }

            _outputWindowPane.OutputString(message + Environment.NewLine);
            _outputWindowPane.Activate(); // Brings this pane into view
        }
    }
}