using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Windows.Forms;
using Microsoft.VisualStudio;
using System.IO;
using System.Reflection;

namespace VSSettingsManager
{
    /// <summary>
    /// Command handler
    /// </summary>
    public sealed class VSSettingsManager
    {
        /// <summary>
        /// Match with symbols in VSCT file.
        /// </summary>
        public static readonly Guid VSSettingsManagerCmdSetGuid = new Guid("cca0811b-addf-4d7b-9dd6-fdb412c44d8a");
        public const int ApplyShortcutsCmdId = 0x0100;
        public const int BackupShortcutsCmdId = 0x0200;
        public const int RestoreShortcutsCmdId = 0x0300;
        public const int ResetShortcutsCmdId = 0x0400;

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static VSSettingsManager Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new VSSettingsManager(package);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VSSettingsManager"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private VSSettingsManager(Package package)
        {
            // Register this command with the Global Command Service
            this.package = package ?? throw new ArgumentNullException("package");

            if (ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                commandService.AddCommand(CreateMenuItem(BackupShortcutsCmdId, this.BackupShortcuts));
                commandService.AddCommand(CreateMenuItem(RestoreShortcutsCmdId, this.RestoreShortcuts));
                commandService.AddCommand(CreateMenuItem(ResetShortcutsCmdId, this.ResetShortcuts));
            }
        }

        private MenuCommand CreateMenuItem(int cmdId, EventHandler menuItemCallback)
        {
            return new MenuCommand(menuItemCallback, new CommandID(VSSettingsManagerCmdSetGuid, cmdId));
        }

        private void BackupShortcuts(object sender, EventArgs e)
        {
            const string Caption = "Backup Shortcuts?";
            const string Text = "Feature not implemented yet.\n" +
                "Go to Tools->Import and Export settings...";
            MessageBox.Show(Text, Caption, MessageBoxButtons.OK);
        }

        private void RestoreShortcuts(object sender, EventArgs e)
        {
            const string Caption = "Restore Shortcuts?";
            const string Text = "Feature not implemented yet.\n" +
                "Go to Tools->Import and Export settings...";
            MessageBox.Show(Text, Caption, MessageBoxButtons.OK);
        }

        private void ResetShortcuts(object sender, EventArgs e)
        {
            const string Caption = "Reset Shortcuts?";
            const string Text = "Feature not implemented yet.\n" +
                "Go to Tools->Import and Export settings...";
            MessageBox.Show(Text, Caption, MessageBoxButtons.OK);
        }

        //------------ Shortcut Settings --------------

        private void ImportUserSettings(string settingsFileName)
        {
            if (ServiceProvider.GetService(typeof(SVsUIShell)) is IVsUIShell shell)
            {
                // import the settings file into Visual Studio
                var asmDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var settingsFilePath = Path.Combine(asmDirectory, settingsFileName);
                var group = VSConstants.CMDSETID.StandardCommandSet2K_guid;

                object arguments = string.Format(CultureInfo.InvariantCulture, "-import:\"{0}\"", settingsFilePath);
                shell.PostExecCommand(ref group, (uint)VSConstants.VSStd2KCmdID.ManageUserSettings, 0, ref arguments);
            }
        }

    }
}