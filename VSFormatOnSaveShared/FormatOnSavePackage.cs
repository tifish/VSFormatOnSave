using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using DefGuidList = Microsoft.VisualStudio.Editor.DefGuidList;
using IServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;
#if true
using Task = System.Threading.Tasks.Task;
using Application = System.Windows.Application;
#endif

namespace Tinyfish.FormatOnSave
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideAutoLoad("{f1536ef8-92ec-443c-9ed7-fdadf150da82}", PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(GuidList.GuidFormatOnSavePkgString)]
    [ProvideOptionPage(typeof(OptionsPage), "Format on Save", "Settings", 0, 0, true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public class FormatOnSavePackage : AsyncPackage
    {
        public DTE2 Dte { get; private set; }
        public OptionsPage OptionsPage { get; private set; }
        private RunningDocumentTable _runningDocumentTable;
        private ServiceProvider _serviceProvider;
        public OleMenuCommandService MenuCommandService { get; private set; }
        private SolutionExplorerContextMenu _solutionExplorerContextMenu;
        public int MajorVersion { get; private set; }
        public int MinorVersion { get; private set; }
        private Events2 _dteEvents;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            Dte = await GetServiceAsync(typeof(SDTE)) as DTE2;
            Assumes.Present(Dte);

            _dteEvents = (Events2)Dte.Events;

            var versionItems = Dte.Version.Split('.');
            if (versionItems.Length == 2)
            {
                MajorVersion = int.Parse(versionItems[0]);
                MinorVersion = int.Parse(versionItems[1]);
            }

            var docTableEventHandler = new VsRunningDocTableEventsHandler(this);

            MenuCommandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;

            // When Initialised asynchronously, the current thread may be a background thread at this point.
            // Do any Initialisation that requires the UI thread after switching to the UI thread.
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            OptionsPage = (OptionsPage)GetDialogPage(typeof(OptionsPage));
            OptionsPage.OnSettingsUpdated += OnSettingsUpdated;

            _serviceProvider = new ServiceProvider((IServiceProvider)Dte);

            _runningDocumentTable = new RunningDocumentTable(this);
            _runningDocumentTable.Advise(docTableEventHandler);

            _solutionExplorerContextMenu = new SolutionExplorerContextMenu(this);

            await EnableDisableFormatOnSaveCommand.InitializeAsync(this);

            UpdateAutoSaveEvents();
        }

        private void OnSettingsUpdated(object sender, EventArgs e)
        {
            UpdateCaptureEvents();
            UpdateAutoSaveEvents();
        }

        #region Auto save events

        private void UpdateAutoSaveEvents()
        {
            if (OptionsPage.EnableAutoSaveOnDeativated)
            {
                if (!_autoSaveEventsRegistered)
                {
                    Application.Current.Deactivated += OnAutoSave;
                    _autoSaveEventsRegistered = true;
                }
            }
            else
            {
                if (_autoSaveEventsRegistered)
                {
                    Application.Current.Deactivated -= OnAutoSave;
                    _autoSaveEventsRegistered = false;
                }
            }
        }

        private bool _autoSaveEventsRegistered = false;

        private void OnAutoSave(object sender, EventArgs e)
        {
            try
            {
                Dte.ExecuteCommand("File.SaveAll");
            }
            catch
            {
                // Ignored
            }
        }

        #endregion

        public void Format(uint docCookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_isSavingWithoutFormatting)
                return;

            var document = FindDocument(docCookie);
            Format(document);
        }

        private Document FindDocument(uint docCookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var documentInfo = _runningDocumentTable.GetDocumentInfo(docCookie);
            var documentPath = documentInfo.Moniker;

            foreach (Document doc in Dte.Documents)
                if (doc.FullName == documentPath)
                    return doc;

            return null;
        }

        private bool _isFormatting;

        public bool Format(Document document)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (document == null || document.Type != "Text" || document.Language == null)
                return false;

            // Edit.FormatDocument command only apply to the active document.
            Document oldActiveDocument = null;
            if (OptionsPage.EnableFormatDocument && document != Dte.ActiveDocument)
            {
                oldActiveDocument = Dte.ActiveDocument;
                document.Activate();
            }

            // Document.Language is unreliable, e.g. "HTMLX" in VS2019 but "Razor" in VS2022. Fallback to extension.
            var ext = Path.GetExtension(document.FullName).ToLower();

            _isFormatting = true;

            try
            {
                // languageOptions.Item("InsertTabs").Value is not affected by the value in .editorconfig, so ignore it.
                //var insertTabs = true;
                //if (document.Language != "Plain Text") // .feature file is "Plain Text", cannot be ignore
                //{
                //    var languageOptions = Dte.Properties["TextEditor", document.Language];
                //    insertTabs = (bool)languageOptions.Item("InsertTabs").Value;
                //}

                var vsTextView = GetIVsTextView(document.FullName);
                if (vsTextView == null)
                    return false;
                var wpfTextView = GetWpfTextView(vsTextView);
                if (wpfTextView == null)
                    return false;

                // In VS2022 .cshtml file, undo cause some error which totally disabled undo function.
                // It seems undo is not necessary in new VS versions.
                //_undoHistoryRegistry.TryGetHistory(wpfTextView.TextBuffer, out var history);
                //using (var undo = history?.CreateTransaction("Format on save"))

                // Keep and restore column later.
                vsTextView.GetCaretPos(out var oldCaretLine, out var oldCaretColumn);
                vsTextView.SetCaretPos(oldCaretLine, 0);

                // Do TabToSpace before FormatDocument, since VS format may break the tab formatting.
                if (OptionsPage.EnableTabToSpace && OptionsPage.AllowDenyTabToSpaceFilter.IsAllowed(document.Name))
                    TabToSpace(wpfTextView, document.TabSize);

                if (OptionsPage.EnableRemoveAndSort && OptionsPage.AllowDenyRemoveAndSortFilter.IsAllowed(document.Name)
                    && ext == ".cs")
                    if (!OptionsPage.EnableSmartRemoveAndSort || !HasIfCompilerDirective(wpfTextView))
                        RemoveAndSort();

                if (OptionsPage.EnableFormatDocument && OptionsPage.AllowDenyFormatDocumentFilter.IsAllowed(document.Name))
                    FormatDocument(ext);

                // Do TabToSpace again after FormatDocument, since VS2017 may stick to tab. Should remove this after VS2017 fix the bug.
                // At 2021.10 the bug has gone. But VS seems to stick to space now, new bug?
                // At 2023.01 when formatting in project/solution, with a .editorconfig set to space,
                // FormatDocument use tab instead. It seems FormatDocument ignore .editorconfig, and only apply VS's settings.
                if (OptionsPage.EnableTabToSpace && OptionsPage.AllowDenyTabToSpaceFilter.IsAllowed(document.Name)
                    && document.Language == "C/C++")
                    TabToSpace(wpfTextView, document.TabSize);

                if (OptionsPage.EnableUnifyLineBreak && OptionsPage.AllowDenyUnifyLineBreakFilter.IsAllowed(document.Name))
                    UnifyLineBreak(wpfTextView, OptionsPage.ForceCRLFFilter.IsAllowed(document.Name));

                if (OptionsPage.EnableUnifyEndOfFile && OptionsPage.AllowDenyUnifyEndOfFileFilter.IsAllowed(document.Name))
                    UnifyEndOfFile(wpfTextView);

                if (OptionsPage.EnableForceUtf8WithBom && OptionsPage.AllowDenyForceUtf8WithBomFilter.IsAllowed(document.Name))
                    ForceUtf8WithBom(wpfTextView);

                // Caret stay in new line but keep old column.
                vsTextView.GetCaretPos(out var newCaretLine, out _);
                vsTextView.SetCaretPos(newCaretLine, oldCaretColumn);
            }
            finally
            {
                oldActiveDocument?.Activate();
                _isFormatting = false;
            }

            return true;
        }

        private static bool HasIfCompilerDirective(ITextView wpfTextView)
        {
            return wpfTextView.TextSnapshot.GetText().Contains("#if");
        }

        private void RemoveAndSort()
        {
            try
            {
                Dte.ExecuteCommand("Edit.RemoveAndSort", "");
            }
            catch (COMException)
            {
            }
        }

        #region Capture delayed Edit.FormatDocument command

        private TextEditorEvents _textEditorEvents;
        private WindowKeyboardHook _keyboardHook;

        private void UpdateCaptureEvents()
        {
            if (MajorVersion < 17)
                return;

            ThreadHelper.ThrowIfNotOnUIThread();

            if (OptionsPage.Enabled && OptionsPage.EnableFormatDocument)
            {
                if (_textEditorEvents == null)
                {
                    _textEditorEvents = _dteEvents.TextEditorEvents;
                    _textEditorEvents.LineChanged += OnLineChanged;
                }

                if (_keyboardHook == null)
                {
                    _keyboardHook = new WindowKeyboardHook();
                    _keyboardHook.OnMessage += OnKeyboardMessage;
                    _keyboardHook.Install();
                }
            }
            else
            {
                if (_textEditorEvents != null)
                {
                    _textEditorEvents.LineChanged -= OnLineChanged;
                    _textEditorEvents = null;
                }

                if (_keyboardHook != null)
                {
                    _keyboardHook.Uninstall();
                    _keyboardHook = null;
                }
            }
        }

        private bool _isSavingWithoutFormatting;

        private void OnLineChanged(TextPoint startPoint, TextPoint endPoint, int hint)
        {
            if (!_isWaitingForDelayedFormatDocumentCommand)
                return;
            if (_isFormatting)
                return;

            // No user typing, but document changed, should be changed by delayed Edit.FormatDocument command.

            // Edit.FormatDocument may trigger multiple modification. Capture the last one.
            ThreadHelper.ThrowIfNotOnUIThread();
            if (startPoint.AbsoluteCharOffset != endPoint.AbsoluteCharOffset)
                return;

            _isWaitingForDelayedFormatDocumentCommand = false;

            // Only process the triggered document
            if (Dte.ActiveDocument == _delayedFormattingDocument)
            {
                // Save without formatting
                _isSavingWithoutFormatting = true;
                try
                {
                    Dte.ExecuteCommand("File.SaveSelectedItems");
                }
                finally
                {
                    _isSavingWithoutFormatting = false;
                }
            }

            _delayedFormattingDocument = null;
        }

        private readonly Keys[] _bypassKeys =
        {
            Keys.Up, Keys.Down, Keys.Left, Keys.Right,
            Keys.Home, Keys.End, Keys.PageUp, Keys.PageDown,
            Keys.Escape, Keys.CapsLock,
            Keys.ControlKey, Keys.ShiftKey, Keys.Alt,
            Keys.F1, Keys.F2, Keys.F3, Keys.F4, Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12,
        };

        private void OnKeyboardMessage(Keys key, bool isPressing)
        {
            // Any user modification will stop waiting for delayed Edit.FormatDocument command.
            // It is not 100% accurrate.
            if (!_isWaitingForDelayedFormatDocumentCommand)
                return;
            if (!isPressing)
                return;
            if (_bypassKeys.Contains(key))
                return;

            _isWaitingForDelayedFormatDocumentCommand = false;
            _delayedFormattingDocument = null;
        }

        private bool _isWaitingForDelayedFormatDocumentCommand;
        private Document _delayedFormattingDocument;

        #endregion

        private void FormatDocument(string ext)
        {
            try
            {
                Dte.ExecuteCommand("Edit.FormatDocument", "");

                // In VS2022, .razor and .cshtml file will delayed Edit.FormatDocument command, which modify document after saving.
                // I try to capture the modification and save again.
                if (MajorVersion >= 17
                    && OptionsPage.DelayedFormatDocumentFilter.IsAllowed(ext))
                {
                    UpdateCaptureEvents();

                    _isWaitingForDelayedFormatDocumentCommand = true;
                    _delayedFormattingDocument = Dte.ActiveDocument;
                }
            }
            catch (COMException)
            {
                // For document.Language == "Plain Text", this always raise an exception.
                // For .feature file of SpecFlow, if nothing to format, exception will be raised.
            }
        }

        private void UnifyLineBreak(ITextView wpfTextView, bool forceCRLF)
        {
            var snapshot = wpfTextView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                string defaultLineBreak;
                if (forceCRLF)
                    defaultLineBreak = "\r\n";
                else
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

        private static void UnifyEndOfFile(ITextView textView)
        {
            var snapshot = textView.TextSnapshot;
            using (var edit = snapshot.TextBuffer.CreateEdit())
            {
                var lineNumber = snapshot.LineCount - 1;
                while (lineNumber >= 0 && snapshot.GetLineFromLineNumber(lineNumber).GetText().Trim() == "")
                    lineNumber--;

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
                // Delete redundant empty lines
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

        private static void ForceUtf8WithBom(ITextView wpfTextView)
        {
            try
            {
                wpfTextView.TextDataModel.DocumentBuffer.Properties.TryGetProperty(typeof(ITextDocument),
                    out ITextDocument textDocument);

                if (textDocument.Encoding.EncodingName != Encoding.UTF8.EncodingName)
                    textDocument.Encoding = Encoding.UTF8;
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private static void RemoveTrailingSpaces(ITextView textView)
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

        private class SpaceStringPool
        {
            private readonly string[] _stringCache = new string[8];

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

        private readonly SpaceStringPool _spaceStringPool = new SpaceStringPool();

        private void TabToSpace(ITextView wpfTextView, int tabSize)
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
                        {
                            positionOffset++;
                        }
                    }
                }

                if (hasModifed)
                    edit.Apply();
            }
        }

        private readonly Regex _cjkRegex = new Regex(
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

        private bool IsCjkCharacter(char character)
        {
            return _cjkRegex.IsMatch(character.ToString());
        }

        private static IWpfTextView GetWpfTextView(IVsTextView vTextView)
        {
            IWpfTextView view = null;
            var userData = (IVsUserData)vTextView;

            if (userData != null)
            {
                var guidViewHost = DefGuidList.guidIWpfTextViewHost;
                userData.GetData(ref guidViewHost, out var holder);
                var viewHost = (IWpfTextViewHost)holder;
                view = viewHost.TextView;
            }

            return view;
        }

        private IVsTextView GetIVsTextView(string filePath)
        {
            return VsShellUtilities.IsDocumentOpen(_serviceProvider, filePath, Guid.Empty, out var uiHierarchy,
                out var itemId, out var windowFrame)
                ? VsShellUtilities.GetTextView(windowFrame)
                : null;
        }

        private IVsOutputWindowPane _outputWindowPane;
        private Guid _outputWindowPaneGuid = new Guid("8AEEC946-659A-4D14-8340-730B2A0FF39C");

        public void OutputString(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_outputWindowPane == null)
            {
                var outWindow = (IVsOutputWindow)GetGlobalService(typeof(SVsOutputWindow));
                outWindow.CreatePane(ref _outputWindowPaneGuid, "VSFormatOnSave", 1, 1);
                outWindow.GetPane(ref _outputWindowPaneGuid, out _outputWindowPane);
            }

            _outputWindowPane.OutputString(message + Environment.NewLine);
            _outputWindowPane.Activate(); // Brings this pane into view
        }
    }
}
