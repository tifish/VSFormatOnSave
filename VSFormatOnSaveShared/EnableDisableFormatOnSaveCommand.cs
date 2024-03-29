﻿using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using Task = System.Threading.Tasks.Task;

namespace Tinyfish.FormatOnSave
{
    /// <summary>
    ///     Command handler
    /// </summary>
    sealed class EnableDisableFormatOnSaveCommand
    {
        /// <summary>
        ///     Command ID.
        /// </summary>
        public const int CommandId = 256;

        /// <summary>
        ///     Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("d787d099-5ebf-4f09-9ef3-10d7cda25bcd");

        /// <summary>
        ///     VS Package that provides this command, not null.
        /// </summary>
        private readonly FormatOnSavePackage package;

        /// <summary>
        ///     Initializes a new instance of the <see cref="EnableDisableFormatOnSaveCommand" /> class.
        ///     Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private EnableDisableFormatOnSaveCommand(FormatOnSavePackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandID);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        ///     Gets the instance of the command.
        /// </summary>
        public static EnableDisableFormatOnSaveCommand Instance { get; private set; }

        /// <summary>
        ///     Gets the service provider from the owner package.
        /// </summary>
        private IAsyncServiceProvider ServiceProvider => package;

        /// <summary>
        ///     Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(FormatOnSavePackage package)
        {
            // Switch to the main thread - the call to AddCommand in EnableDisableFormatOnSaveCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new EnableDisableFormatOnSaveCommand(package, commandService);
        }

        /// <summary>
        ///     This function is the callback used to execute the command when the menu item is clicked.
        ///     See the constructor to see how the menu item is associated with this function using
        ///     OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            var command = sender as OleMenuCommand;
            if (command == null)
                return;
            // User click menu command before extension loading.
            if (command.Text == "")
                return;

            package.OptionsPage.Enabled = !package.OptionsPage.Enabled;
            package.OptionsPage.SaveSettingsToStorage();
            package.OptionsPage.UpdateSettings();
        }

        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            var command = sender as OleMenuCommand;
            if (command == null)
                return;

            command.Checked = package.OptionsPage.Enabled;
            command.Text = package.OptionsPage.Enabled ? "Format on Save is on, click to turn off" : "Format on Save is off, click to turn on";
        }
    }
}
