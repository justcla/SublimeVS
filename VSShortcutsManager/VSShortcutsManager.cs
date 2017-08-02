using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Windows.Forms;
using Microsoft.VisualStudio;
using System.IO;
using System.Reflection;

namespace VSShortcutsManager
{
    /// <summary>
    /// Command handler
    /// </summary>
    public sealed class VSShortcutsManager
    {
        /// <summary>
        /// Match with symbols in VSCT file.
        /// </summary>
        public static readonly Guid VSShortcutsManagerCmdSetGuid = new Guid("cca0811b-addf-4d7b-9dd6-fdb412c44d8a");
        public const int BackupShortcutsCmdId = 0x0200;
        public const int RestoreShortcutsCmdId = 0x0300;
        public const int ResetShortcutsCmdId = 0x0400;

        private const string BACKUP_FILE_PATH = "BackupFilePath";
        private const string MSG_CAPTION_RESTORE = "Restore Keyboard Shortcuts";
        private const string MSG_CAPTION_BACKUP = "Backup Keyboard Shortcuts";
        private const string MSG_CAPTION_RESET = "Reset Keyboard Shortcuts";


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
        public static VSShortcutsManager Instance
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
            Instance = new VSShortcutsManager(package);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VSShortcutsManager"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private VSShortcutsManager(Package package)
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
            return new MenuCommand(menuItemCallback, new CommandID(VSShortcutsManagerCmdSetGuid, cmdId));
        }

        private void BackupShortcuts(object sender, EventArgs e)
        {
            ExecuteBackupShortcuts();
        }

        private void RestoreShortcuts(object sender, EventArgs e)
        {
            ExecuteRestoreShortcuts();
        }

        private void ResetShortcuts(object sender, EventArgs e)
        {
            const string Text = "Feature not implemented yet.\n\n" +
                "Look for Reset under Tools->Options; Environment->Keyboard";
            MessageBox.Show(Text, MSG_CAPTION_RESET, MessageBoxButtons.OK);
        }

        //-------- Backup Shortcuts --------

        public void ExecuteBackupShortcuts()
        {
            // Confirm Backup operation
            if (MessageBox.Show("Backup current keyboard shortcuts?", MSG_CAPTION_BACKUP, MessageBoxButtons.OKCancel) != DialogResult.OK)
            {
                return;
            }

            IVsProfileDataManager vsProfileDataManager = (IVsProfileDataManager)ServiceProvider.GetService(typeof(SVsProfileDataManager));

            // Get the filename where the vssettings file will be saved
            string backupFilePath = GetExportFilePath(vsProfileDataManager);

            // Do the export
            IVsProfileSettingsTree keyboardOnlyExportSettings = GetKeyboardOnlyExportSettings(vsProfileDataManager);
            int result = vsProfileDataManager.ExportSettings(backupFilePath, keyboardOnlyExportSettings, out IVsSettingsErrorInformation errorInfo);
            if (result != VSConstants.S_OK)
            {
                // Something went wrong. TODO: Handle error.
            }

            // Save Backup file path to SettingsManager
            SaveBackupFilePath(backupFilePath);

            // Report success
            string Text = $"Your keyboard shortcuts have been backed up to the following file:\n\n{backupFilePath}";
            MessageBox.Show(Text, MSG_CAPTION_BACKUP, MessageBoxButtons.OK);
        }

        private static string GetExportFilePath(IVsProfileDataManager vsProfileDataManager)
        {
            vsProfileDataManager.GetUniqueExportFileName((uint)__VSPROFILEGETFILENAME.PGFN_SAVECURRENT, out string exportFilePath);
            return exportFilePath;
        }

        private static IVsProfileSettingsTree GetKeyboardOnlyExportSettings(IVsProfileDataManager vsProfileDataManager)
        {
            vsProfileDataManager.GetSettingsForExport(out IVsProfileSettingsTree profileSettingsTree);
            // Disable all settings for export
            profileSettingsTree.SetEnabled(0, 1);
            // Enable Keyboard settings for export
            profileSettingsTree.FindChildTree("Environment_Group\\Environment_KeyBindings", out IVsProfileSettingsTree keyboardSettingsTree);
            if (keyboardSettingsTree != null)
            {
                int enabledValue = 1;  // true
                int setChildren = 0;  // true  (redundant)
                keyboardSettingsTree.SetEnabled(enabledValue, setChildren);
            }
            return profileSettingsTree;
        }

        //------------ Restore Shortcuts --------------

        public void ExecuteRestoreShortcuts()
        {
            string backupFilePath = GetSavedBackupFilePath();
            if (String.IsNullOrEmpty(backupFilePath))
            {
                MessageBox.Show("Unable to restore keyboard shortcuts.\n\nReason: No known backup file has been created.", MSG_CAPTION_RESTORE);
                return;
            }

            string text = $"Restore keyboard shortcuts from the last backup?\n" +
                $"\nLast backup location:\n{backupFilePath}";
            if (MessageBox.Show(text, MSG_CAPTION_RESTORE, MessageBoxButtons.OKCancel) != DialogResult.OK)
            {
                return;
            }

            ImportSettingsFromFilePath(backupFilePath);

            MessageBox.Show("Keyboard shortcuts successfully restored.", MSG_CAPTION_RESTORE);
        }

        private void ImportUserSettings(string settingsFileName)
        {
            // import the settings file into Visual Studio
            var asmDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var settingsFilePath = Path.Combine(asmDirectory, settingsFileName);
            ImportSettingsFromFilePath(settingsFilePath);
        }

        private void ImportSettingsFromFilePath(string settingsFilePath)
        {
            var group = VSConstants.CMDSETID.StandardCommandSet2K_guid;
            if (ServiceProvider.GetService(typeof(SVsUIShell)) is IVsUIShell shell)
            {
                object arguments = string.Format(CultureInfo.InvariantCulture, "-import:\"{0}\"", settingsFilePath);
                // NOTE: Call to PostExecCommand could fail. Callers should consider catching the exception. Otherwise, UI will show the error in a messagebox.
                shell.PostExecCommand(ref group, (uint)VSConstants.VSStd2KCmdID.ManageUserSettings, 0, ref arguments);
            }
        }

        //---------- Helper methods -------------------

        private static void SaveBackupFilePath(string exportFilePath)
        {
            try
            {
                VSShortcutsManagerPackage.SettingsManager.SetValueAsync(BACKUP_FILE_PATH, exportFilePath, isMachineLocal: true);
            }
            catch (Exception e)
            {
                // Unable to save backup file location. TODO: Handle error.
            }
        }

        private string GetSavedBackupFilePath()
        {
            VSShortcutsManagerPackage.SettingsManager.TryGetValue(BACKUP_FILE_PATH, out string backupFilePath);
            return backupFilePath;
        }

    }
}