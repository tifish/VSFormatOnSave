using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace Tinyfish.FormatOnSave
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("1D9ECCF3-5D2F-4112-9B25-264596873DC9")]
    public class OptionsPage : DialogPage
    {
        [Category("All")]
        [DisplayName("Enable FormatOnSave")]
        [Description("Enable all FormatOnSave features.")]
        public bool Enabled { get; set; } = true;


        [Category("Auto Save")]
        [DisplayName("Enable AutoSaveOnDeativated")]
        [Description("Enable AutoSaveOnDeativated, save files when Visual Studio deactivated.")]
        public bool EnableAutoSaveOnDeativated { get; set; } = true;


        public enum LineBreakStyle
        {
            Unix = 0,
            Windows = 1,
        }

        [Category("Remove and sort for C#")]
        [DisplayName("Enable RemoveAndSort")]
        [Description("Enable RemoveAndSort, remove and sort using statements. Only apply to .cs file.")]
        public bool EnableRemoveAndSort { get; set; } = false;

        [Category("Remove and sort for C#")]
        [DisplayName("Enable SmartRemoveAndSort")]
        [Description("SmartRemoveAndSort apply RemoveAndSort to .cs without #if only. Remove and sort must be enabled first.")]
        public bool EnableSmartRemoveAndSort { get; set; } = true;

        [Category("Remove and sort for C#")]
        [DisplayName("Allowed extensions for RemoveAndSort only")]
        [Description("Allowed extensions for RemoveAndSort only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowRemoveAndSortExtensions { get; set; } = string.Empty;

        [Category("Remove and sort for C#")]
        [DisplayName("Denied extensions for RemoveAndSort only")]
        [Description("Denied extensions for RemoveAndSort only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyRemoveAndSortExtensions { get; set; } = string.Empty;

        public AllowDenyDocumentFilter AllowDenyRemoveAndSortFilter;


        [Category("Format document")]
        [DisplayName("Enable FormatDocument")]
        [Description("Enable FormatDocument, automatically format document when saving.")]
        public bool EnableFormatDocument { get; set; } = true;

        [Category("Format document")]
        [DisplayName("Allowed extensions for FormatDocument only")]
        [Description("Allowed extensions for FormatDocument only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowFormatDocumentExtentions { get; set; } = string.Empty;

        [Category("Format document")]
        [DisplayName("Denied extensions for FormatDocument only")]
        [Description("Denied extensions for FormatDocument only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyFormatDocumentExtentions { get; set; } = string.Empty;

        public AllowDenyDocumentFilter AllowDenyFormatDocumentFilter;

        [Category("Format document")]
        [DisplayName("Extensions cause delayed FormatDocument")]
        [Description("Extensions cause delayed FormatDocument, which modify file after saving. Space separated list. For example: .razor .cshtml")]
        public string DelayedFormatDocumentExtentions { get; set; } = ".razor .cshtml";

        public AllowDenyDocumentFilter DelayedFormatDocumentFilter;


        [Category("Line break")]
        [DisplayName("Enable UnifyLineBreak")]
        [Description("Enable UnifyLineBreak.")]
        public bool EnableUnifyLineBreak { get; set; } = false;

        [Category("Line break")]
        [DisplayName("Line break style")]
        [Description("Line break style.")]
        public LineBreakStyle LineBreak { get; set; } = LineBreakStyle.Unix;

        [Category("Line break")]
        [DisplayName("Allowed extensions for UnifyLineBreak only")]
        [Description("Allowed extensions for UnifyLineBreak only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowUnifyLineBreakExtensions { get; set; } = string.Empty;

        [Category("Line break")]
        [DisplayName("Denied extensions for UnifyLineBreak only")]
        [Description("Denied extensions for UnifyLineBreak only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyUnifyLineBreakExtensions { get; set; } = string.Empty;

        public AllowDenyDocumentFilter AllowDenyUnifyLineBreakFilter;

        [Category("Line break")]
        [DisplayName("Force CRLF file extensions")]
        [Description("Force CRLF file extensions. Some files like .aspx must be CRLF, or Visual Studio format will produce weird empty lines. Space separated list. For example: .cs .html .cshtml .vb")]
        public string ForceCRLFExtensions { get; set; } = ".aspx";

        public AllowDenyDocumentFilter ForceCRLFFilter;


        [Category("End of file")]
        [DisplayName("Enable UnifyEndOfFile")]
        [Description("Enable UnifyEndOfFile, ensure file end with one return.")]
        public bool EnableUnifyEndOfFile { get; set; } = false;

        [Category("End of file")]
        [DisplayName("Allowed extensions for UnifyEndOfFile only")]
        [Description("Allowed extensions for UnifyEndOfFile only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowUnifyEndOfFileExtensions { get; set; } = string.Empty;

        [Category("End of file")]
        [DisplayName("Denied extensions for UnifyEndOfFile only")]
        [Description("Denied extensions for UnifyEndOfFile only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyUnifyEndOfFileExtensions { get; set; } = string.Empty;

        public AllowDenyDocumentFilter AllowDenyUnifyEndOfFileFilter;


        [Obsolete]
        [Category("Others (deprecated)")]
        [DisplayName("Allowed extensions for all except FormatDocument")]
        [Description("Allowed extensions for all except FormatDocument. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowExtentions { get; set; } = string.Empty;

        [Obsolete]
        [Category("Others (deprecated)")]
        [DisplayName("Denied extensions for all except FormatDocument")]
        [Description(
            "Denied extensions for all except FormatDocument. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyExtentions { get; set; } = string.Empty;


        [Category("Tab to space")]
        [DisplayName("Enable TabToSpace")]
        [Description("Enable TabToSpace. Depends on tabs options for the type of file.")]
        public bool EnableTabToSpace { get; set; } = false;

        [Category("Tab to space")]
        [DisplayName("Allowed extensions for TabToSpace only")]
        [Description("Allowed extensions for TabToSpace only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowTabToSpaceExtensions { get; set; } = string.Empty;

        [Category("Tab to space")]
        [DisplayName("Denied extensions for TabToSpace only")]
        [Description("Denied extensions for TabToSpace only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyTabToSpaceExtensions { get; set; } = string.Empty;

        public AllowDenyDocumentFilter AllowDenyTabToSpaceFilter;


        [Category("UTF8")]
        [DisplayName("Enable ForceUtf8WithBom")]
        [Description("Enable ForceUtf8WithBom, force file encoding to UTF8 with BOM.")]
        public bool EnableForceUtf8WithBom { get; set; } = false;

        [Category("UTF8")]
        [DisplayName("Allowed extensions for ForceUtf8WithBom only")]
        [Description("Allowed extensions for ForceUtf8WithBom only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowForceUtf8WithBomExtentions { get; set; } = string.Empty;

        [Category("UTF8")]
        [DisplayName("Denied extensions for ForceUtf8WithBom only")]
        [Description("Denied extensions for ForceUtf8WithBom only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyForceUtf8WithBomExtentions { get; set; } = string.Empty;

        public AllowDenyDocumentFilter AllowDenyForceUtf8WithBomFilter;

        public event EventHandler OnSettingsUpdated;

        public void UpdateSettings()
        {
            AllowDenyRemoveAndSortFilter = new AllowDenyDocumentFilter(
                AllowRemoveAndSortExtensions.Split(' '), DenyRemoveAndSortExtensions.Split(' '));

            AllowDenyFormatDocumentFilter = new AllowDenyDocumentFilter(
                AllowFormatDocumentExtentions.Split(' '), DenyFormatDocumentExtentions.Split(' '));

            DelayedFormatDocumentFilter = new AllowDenyDocumentFilter(DelayedFormatDocumentExtentions.Split(' '), null);

            AllowDenyUnifyLineBreakFilter = new AllowDenyDocumentFilter(
                AllowUnifyLineBreakExtensions.Split(' '), DenyUnifyLineBreakExtensions.Split(' '));

            if (string.IsNullOrWhiteSpace(ForceCRLFExtensions))
                ForceCRLFExtensions = ".aspx";
            ForceCRLFFilter = new AllowDenyDocumentFilter(ForceCRLFExtensions.Split(' '), null);

            AllowDenyUnifyEndOfFileFilter = new AllowDenyDocumentFilter(
                AllowUnifyEndOfFileExtensions.Split(' '), DenyUnifyEndOfFileExtensions.Split(' '));

            AllowDenyTabToSpaceFilter = new AllowDenyDocumentFilter(
                AllowTabToSpaceExtensions.Split(' '), DenyTabToSpaceExtensions.Split(' '));

            AllowDenyForceUtf8WithBomFilter = new AllowDenyDocumentFilter(
                AllowForceUtf8WithBomExtentions.Split(' '), DenyForceUtf8WithBomExtentions.Split(' '));

            OnSettingsUpdated?.Invoke(this, null);
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
