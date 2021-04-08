using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;

namespace Tinyfish.FormatOnSave
{
    class SolutionExplorerContextMenu
    {
        readonly FormatOnSavePackage _package;

        public SolutionExplorerContextMenu(FormatOnSavePackage package)
        {
            _package = package;

            var mcs = _package.MenuCommandService;

            var menuCommandId = new CommandID(GuidList.GuidFormatOnSaveCmdSetFile, (int)PkgCmdIdList.CmdIdFormatOnSaveFile);
            var menuItem = new MenuCommand(FormatOnSaveInFileEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(GuidList.GuidFormatOnSaveCmdSetFolder, (int)PkgCmdIdList.CmdIdFormatOnSaveFolder);
            menuItem = new MenuCommand(FormatOnSaveInFolderEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(GuidList.GuidFormatOnSaveCmdSetProject, (int)PkgCmdIdList.CmdIdFormatOnSaveProject);
            menuItem = new MenuCommand(FormatOnSaveInProjectEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(GuidList.GuidFormatOnSaveCmdSetMultipleItems, (int)PkgCmdIdList.CmdIdFormatOnSaveMultipleItems);
            menuItem = new MenuCommand(FormatOnSaveInProjectEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(GuidList.GuidFormatOnSaveCmdSetSolution, (int)PkgCmdIdList.CmdIdFormatOnSaveSolution);
            menuItem = new MenuCommand(FormatOnSaveInSolutionEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);

            menuCommandId = new CommandID(GuidList.GuidFormatOnSaveCmdSetSolutionFolder, (int)PkgCmdIdList.CmdIdFormatOnSaveSolutionFolder);
            menuItem = new MenuCommand(FormatOnSaveInSolutionFolderEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);
        }

        void FormatOnSaveInFileEventHandler(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            FormatSelectedItems();
        }

        void FormatOnSaveInFolderEventHandler(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            FormatSelectedItems();
        }

        void FormatOnSaveInProjectEventHandler(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            FormatSelectedItems();
        }

        void FormatOnSaveInSolutionEventHandler(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            FormatSelectedItems();
        }

        void FormatOnSaveInSolutionFolderEventHandler(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            FormatSelectedItems();
        }

        void FormatSelectedItems()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (UIHierarchyItem selectedItem in (object[])_package.Dte.ToolWindows.SolutionExplorer.SelectedItems)
            {
                FormatItem(selectedItem.Object);
            }
        }

        void FormatItem(object item)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            switch (item)
            {
                case Solution solution:
                    {
                        foreach (Project subProject in solution.Projects)
                        {
                            FormatItem(subProject);
                        }

                        return;
                    }

                case Project project:
                    {
                        if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                        {
                            foreach (ProjectItem projectSubItem in project.ProjectItems)
                            {
                                FormatItem(projectSubItem.SubProject);
                            }
                        }
                        else
                        {
                            foreach (ProjectItem projectSubItem in project.ProjectItems)
                            {
                                FormatItem(projectSubItem);
                            }
                        }

                        return;
                    }

                case ProjectItem projectItem when projectItem.ProjectItems != null && projectItem.ProjectItems.Count > 0:
                    {
                        foreach (ProjectItem subProjectItem in projectItem.ProjectItems)
                        {
                            FormatItem(subProjectItem);
                        }

                        break;
                    }

                case ProjectItem projectItem:
                    FormatProjectItem(projectItem);
                    break;
            }
        }

        void FormatProjectItem(ProjectItem item)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!_package.OptionsPage.AllowDenyFilter.IsAllowed(item.Name))
                return;

            Window documentWindow = null;
            try
            {
                if (!item.IsOpen)
                {
                    documentWindow = item.Open();
                    if (documentWindow == null)
                        return;
                }

                if (_package.Format(item.Document))
                    item.Document.Save();
            }
            catch (COMException)
            {
                _package.OutputString($"Failed to process {item.Name}.");
            }
            finally
            {
                documentWindow?.Close();
            }
        }
    }
}
