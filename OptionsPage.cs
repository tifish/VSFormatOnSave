using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Tinyfish.FormatOnSave
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("1D9ECCF3-5D2F-4112-9B25-264596873DC9")]
    public class OptionsPage : DialogPage
    {
        public enum LineBreakStyle
        {
            Unix = 0,
            Windows = 1,
        };

        [Category("On Save")]
        [Description("Enable remove and sort on save, only apply to .cs file.")]
        public bool EnableRemoveAndSort { get; set; } = true;

        [Category("On Save")]
        [Description("Apply remove and sort to .cs without #if. Remove and sort must be enabled first.")]
        public bool EnableSmartRemoveAndSort { get; set; } = true;

        [Category("On Save")]
        [Description("Enable format document on save.")]
        public bool EnableFormatDocument { get; set; } = true;

        [Category("On Save")]
        [Description("Allow extentions for FormatDocument only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowFormatDocumentExtentions { get; set; } = string.Empty;

        [Category("On Save")]
        [Description("Deny extentions for FormatDocument only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyFormatDocumentExtentions { get; set; } = string.Empty;

        [Category("On Save")]
        [Description("Enable unify line break on save.")]
        public bool EnableUnifyLineBreak { get; set; } = true;

        [Category("On Save")]
        [Description("Allow extentions for all except FormatDocument. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowExtentions { get; set; } = string.Empty;

        [Category("On Save")]
        [Description(
            "Deny extentions for all except FormatDocument. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyExtentions { get; set; } = string.Empty;

        [Category("On Save")]
        [Description("Line break style.")]
        public LineBreakStyle LineBreak { get; set; } = LineBreakStyle.Unix;

        [Category("On Save")]
        [Description("Enable unify end of file to one empty line on save.")]
        public bool EnableUnifyEndOfFile { get; set; } = true;

        [Category("On Save")]
        [Description("Enable tab to space on save. Depends on tabs options for the type of file.")]
        public bool EnableTabToSpace { get; set; } = true;

        [Category("On Save")]
        [Description("Enable force file encoding to UTF8 with BOM on save.")]
        public bool EnableForceUtf8WithBom { get; set; } = true;

        [Category("On Save")]
        [Description("Allow extentions for ForceUtf8WithBom only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowForceUtf8WithBomExtentions { get; set; } = string.Empty;

        [Category("On Save")]
        [Description("Deny extentions for ForceUtf8WithBom only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyForceUtf8WithBomExtentions { get; set; } = string.Empty;

        [Category("On Save")]
        [Description("Enable remove trailing spaces. It is mostly for Visual Sutdio 2012, which won't remove trailing spaces when formatting. In higher version than 2012, this will do nothing when FormatDocument is enabled.")]
        public bool EnableRemoveTrailingSpaces { get; set; } = true;

        public AllowDenyDocumentFilter AllowDenyFormatDocumentFilter;
        public AllowDenyDocumentFilter AllowDenyForceUtf8WithBomFilter;
        public AllowDenyDocumentFilter AllowDenyFilter;

        void UpdateSettings()
        {
            AllowDenyFormatDocumentFilter = new AllowDenyDocumentFilter(
                AllowFormatDocumentExtentions.Split(' '), DenyFormatDocumentExtentions.Split(' '));

            AllowDenyForceUtf8WithBomFilter = new AllowDenyDocumentFilter(
                AllowForceUtf8WithBomExtentions.Split(' '), DenyForceUtf8WithBomExtentions.Split(' '));

            AllowDenyFilter = new AllowDenyDocumentFilter(
                AllowExtentions.Split(' '), DenyExtentions.Split(' '));
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            base.OnApply(e);
            UpdateSettings();
        }

        public override void LoadSettingsFromStorage()
        {
            base.LoadSettingsFromStorage();
            UpdateSettings();
        }
    }
}
