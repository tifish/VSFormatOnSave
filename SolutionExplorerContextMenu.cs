using EnvDTE;
using System;
using System.ComponentModel.Design;

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

            menuCommandId = new CommandID(GuidList.GuidFormatOnSaveCmdSetSolution, (int)PkgCmdIdList.CmdIdFormatOnSaveSolution);
            menuItem = new MenuCommand(FormatOnSaveInSolutionEventHandler, menuCommandId)
            {
                Visible = true,
                Enabled = true
            };
            mcs.AddCommand(menuItem);
        }

        void FormatOnSaveInFileEventHandler(object sender, EventArgs e)
        {
            foreach (SelectedItem selectedItem in _package.Dte.SelectedItems)
            {
                var projectItem = selectedItem.ProjectItem;
                FormatOnSaveInProjectItem(projectItem);
            }
        }

        void FormatOnSaveInFolderEventHandler(object sender, EventArgs e)
        {
            foreach (SelectedItem selectedItem in _package.Dte.SelectedItems)
            {
                var projectItem = selectedItem.ProjectItem;
                FormatOnSaveInProjectItems(projectItem.ProjectItems);
            }
        }

        void FormatOnSaveInProjectEventHandler(object sender, EventArgs e)
        {
            foreach (SelectedItem selectedItem in _package.Dte.SelectedItems)
            {
                var selectedProject = selectedItem.Project;
                FormatOnSaveInProjectItems(selectedProject.ProjectItems);
            }
        }

        void FormatOnSaveInSolutionEventHandler(object sender, EventArgs e)
        {
            var currentSolution = _package.Dte.Solution;
            foreach (Project project in currentSolution)
            {
                FormatOnSaveInProjectItems(project.ProjectItems);
            }
        }

        void FormatOnSaveInProjectItems(ProjectItems projectItems)
        {
            foreach (ProjectItem item in projectItems)
            {
                if (item.ProjectItems != null && item.ProjectItems.Count > 0)
                {
                    FormatOnSaveInProjectItems(item.ProjectItems);
                }
                else
                {
                    FormatOnSaveInProjectItem(item);
                }
            }
        }

        void FormatOnSaveInProjectItem(ProjectItem item)
        {
            if (!_package.OptionsPage.AllowDenyFilter.IsAllowed(item.Name))
            {
                return;
            }

            Window documentWindow = null;
            if (!item.IsOpen)
            {
                documentWindow = item.Open();
            }

            if (_package.Format(item.Document))
            {
                item.Save();
            }

            documentWindow?.Close();
        }
    }
}
