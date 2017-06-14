using System;

namespace Tinyfish.FormatOnSave
{
    static class GuidList
    {
        public const string GuidFormatOnSavePkgString = "46644c38-fb23-4f71-bc73-d8673b754a1e";

        public const string GuidFormatOnSaveCmdSetStringFile = "e87176c7-5748-4cef-8933-ce9be1c96113";
        public const string GuidFormatOnSaveCmdSetStringFolder = "3ab3c5e3-d02f-43f0-a482-7def23f76e6e";
        public const string GuidFormatOnSaveCmdSetStringProject = "b6afe42c-8081-4381-a0f7-7668bdbb9562";
        public const string GuidFormatOnSaveCmdSetStringSolution = "d876ec3f-7364-4f0f-be30-6b48ec10d9ba";
        public const string GuidFormatOnSaveCmdSetStringSolutionFolder = "184023d6-1c7f-4d8b-a823-2b205fac7308";

        public static readonly Guid GuidFormatOnSaveCmdSetFile = new Guid(GuidFormatOnSaveCmdSetStringFile);
        public static readonly Guid GuidFormatOnSaveCmdSetFolder = new Guid(GuidFormatOnSaveCmdSetStringFolder);
        public static readonly Guid GuidFormatOnSaveCmdSetProject = new Guid(GuidFormatOnSaveCmdSetStringProject);
        public static readonly Guid GuidFormatOnSaveCmdSetSolution = new Guid(GuidFormatOnSaveCmdSetStringSolution);
        public static readonly Guid GuidFormatOnSaveCmdSetSolutionFolder = new Guid(GuidFormatOnSaveCmdSetStringSolutionFolder);
    };
}
