using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TextManager.Interop;
using System.Runtime.InteropServices;

namespace Tinyfish.FormatOnSave
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideAutoLoad("{f1536ef8-92ec-443c-9ed7-fdadf150da82}")] //To set the UI context to autoload a VSPackage
    [Guid(GuidList.GuidFormatOnSavePkgString)]
    [ProvideOptionPage(typeof(OptionsPage), "Format on Save", "Settings", 0, 0, true)]
    public class FormatOnSavePackage : Package
    {
        protected override void Initialize()
        {
            var dte = (DTE)GetService(typeof(DTE));
            var textManager = (IVsTextManager)GetService(typeof(SVsTextManager));
            var runningDocumentTable = new RunningDocumentTable(this);
            var options = (OptionsPage)GetDialogPage(typeof(OptionsPage));
            var plugin = new FormatOnSaveService(dte, runningDocumentTable, textManager, options);
            runningDocumentTable.Advise(plugin);

            base.Initialize();
        }
    }
}
