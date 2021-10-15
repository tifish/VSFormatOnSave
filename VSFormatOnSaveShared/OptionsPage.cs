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
        [DisplayName("Enable format on save")]
        [Description("Enable format on save.")]
        public bool Enabled { get; set; } = true;

        public enum LineBreakStyle
        {
            Unix = 0,
            Windows = 1,
        }

        [Category("Remove and sort for C#")]
        [DisplayName("Enable remove and sort")]
        [Description("Enable remove and sort, only apply to .cs file.")]
        public bool EnableRemoveAndSort { get; set; } = false;

        [Category("Remove and sort for C#")]
        [DisplayName("Enable smart remove and sort")]
        [Description("Apply remove and sort to .cs without #if. Remove and sort must be enabled first.")]
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
        [DisplayName("Enable format document")]
        [Description("Enable format document.")]
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


        [Category("Line break")]
        [DisplayName("Enable unify line break")]
        [Description("Enable unify line break.")]
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
        [DisplayName("Denied extensions for LineBreak only")]
        [Description("Denied extensions for LineBreak only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyUnifyLineBreakExtensions { get; set; } = string.Empty;

        public AllowDenyDocumentFilter AllowDenyUnifyLineBreakFilter;


        [Category("End of file")]
        [DisplayName("Enable unify end of file")]
        [Description("Enable unify end of file to one empty line.")]
        public bool EnableUnifyEndOfFile { get; set; } = true;

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
        [DisplayName("Enable tab to space")]
        [Description("Enable tab to space. Depends on tabs options for the type of file.")]
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
        [DisplayName("Enable force file encoding to UTF8 with BOM")]
        [Description("Enable force file encoding to UTF8 with BOM.")]
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


        void UpdateSettings()
        {
            AllowDenyRemoveAndSortFilter = new AllowDenyDocumentFilter(
                AllowRemoveAndSortExtensions.Split(' '), DenyRemoveAndSortExtensions.Split(' '));

            AllowDenyFormatDocumentFilter = new AllowDenyDocumentFilter(
                AllowFormatDocumentExtentions.Split(' '), DenyFormatDocumentExtentions.Split(' '));

            AllowDenyUnifyLineBreakFilter = new AllowDenyDocumentFilter(
                AllowUnifyLineBreakExtensions.Split(' '), DenyUnifyLineBreakExtensions.Split(' '));

            AllowDenyUnifyEndOfFileFilter = new AllowDenyDocumentFilter(
                AllowUnifyEndOfFileExtensions.Split(' '), DenyUnifyEndOfFileExtensions.Split(' '));

            AllowDenyTabToSpaceFilter = new AllowDenyDocumentFilter(
                AllowTabToSpaceExtensions.Split(' '), DenyTabToSpaceExtensions.Split(' '));

            AllowDenyForceUtf8WithBomFilter = new AllowDenyDocumentFilter(
                AllowForceUtf8WithBomExtentions.Split(' '), DenyForceUtf8WithBomExtentions.Split(' '));
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