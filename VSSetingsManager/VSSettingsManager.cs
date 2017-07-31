using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Windows.Forms;
using Microsoft.VisualStudio;
using System.IO;
using System.Reflection;
using EnvDTE80;
using EnvDTE;
using System.Diagnostics;

namespace VSSetingsManager
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class VSSettingsManager
    {
        /// <summary>
        /// Match with symbols in VSCT file.
        /// </summary>
        public static readonly Guid VSSettingsManagerCmdSetGuid = new Guid("cca0811b-addf-4d7b-9dd6-fdb412c44d8a");
        public const int ApplyShortcutsCmdId = 0x0100;
        public const int BackupShortcutsCmdId = 0x0200;
        public const int RestoreShortcutsCmdId = 0x0300;
        public const int ResetShortcutsCmdId = 0x0400;

        private const string SublimeSettingsFileName = @"Shortcuts\SublimeShortcuts.vssettings";

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
                commandService.AddCommand(CreateMenuItem(ApplyShortcutsCmdId, this.ApplySettings));
                commandService.AddCommand(CreateMenuItem(BackupShortcutsCmdId, this.BackupShortcuts));
                commandService.AddCommand(CreateMenuItem(RestoreShortcutsCmdId, this.RestoreShortcuts));
                commandService.AddCommand(CreateMenuItem(ResetShortcutsCmdId, this.ResetShortcuts));
            }
        }

        private MenuCommand CreateMenuItem(int cmdId, EventHandler menuItemCallback)
        {
            return new MenuCommand(menuItemCallback, new CommandID(VSSettingsManagerCmdSetGuid, cmdId));
        }

        private void ApplySettings(object sender, EventArgs e)
        {
            //ShowDefaultAlert();
            ApplySublimeShortcuts();
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

        public void ApplySublimeShortcuts()
        {
            //Ask user if they want to apply Sublime settings
            if (ConfirmApplyShortcuts())
            {
                ImportUserSettings(SublimeSettingsFileName);
            }
        }

        private static bool ConfirmApplyShortcuts()
        {
            const string title = "Extensions for Accessibility in Visual Studio";
            const string message =
                "Congratulations! Sublime VS Extension has been installed.\n" +
                "\n" +
                "Apply the following settings:\n" +
                "- Turn on Map Mode Scrollbar (Wide)\n" +
                "- Apply Sublime Text keyboard shortcuts\n" +
                "\n" +
                "Note: You can modify these settings later in Tools->Options";
            return MessageBox.Show(message, title, MessageBoxButtons.OKCancel) == DialogResult.OK;
        }

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

        //------------ MiniMap settings ------------

        public void ApplyMiniMap()
        {
            // Confirm with user before applying MiniMap
            if (ConfirmApplyMinimap())
            {
                ApplyMapModeScrollbar();
            }

        }

        private static bool ConfirmApplyMinimap()
        {
            const string title = "Extensions for Accessibility in Visual Studio";
            const string message =
                "Apply the following settings:\n" +
                "- Turn on Map Mode Scrollbar (Wide)\n" +
                "\n" +
                "Note: You can modify these settings later in Tools->Options";
            return MessageBox.Show(message, title, MessageBoxButtons.OKCancel) == DialogResult.OK;
        }

        private void ApplyMapModeScrollbar()
        {
            try
            {
                var dte2 = (DTE2)(ServiceProvider.GetService(typeof(DTE)));
                UpdateSetting(dte2, "TextEditor", "AllLanguages", "UseMapMode", true);
                UpdateSetting(dte2, "TextEditor", "AllLanguages", "OverviewWidth", (short)83);
            }
            catch (Exception e)
            {
                Debug.Print(e.Message);
            }
        }


        private static void UpdateSetting(DTE2 dte2, string category, string page, string settingName, object value)
        {
            // Example: dte2.Properties["TextEditor", "General"].Item("TrackChanges").Value = true;
            dte2.Properties[category, page].Item(settingName).Value = value;
        }

    }
}