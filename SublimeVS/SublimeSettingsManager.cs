using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Globalization;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using System.IO;
using System.Reflection;
using System.ComponentModel.Design;
using System.Threading.Tasks;

namespace SublimeVS
{
    class SublimeSettingsManager
    {
        /// <summary>
        /// Match with symbols in VSCT file.
        /// </summary>
        public static readonly Guid SublimeVSCmdSetGuid = new Guid("741d3ce4-bc84-49ae-8da8-48c574b99dfb");
        public const int ApplySettingsCmdId = 0x0100;

        private const string SublimeSettingsFileName = @"Shortcuts\SublimeShortcuts.vssettings";

        private readonly AsyncPackage package;

        private IAsyncServiceProvider AsyncServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public static SublimeSettingsManager Instance
        {
            get;
            private set;
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            Instance = new SublimeSettingsManager(package);
            await Instance.InitializeAsync();
        }

        private SublimeSettingsManager(AsyncPackage package)
        {
            // Register this command with the Global Command Service
            this.package = package ?? throw new ArgumentNullException("package");
        }

        public async Task InitializeAsync()
        {

            if (await AsyncServiceProvider.GetServiceAsync(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                // Switch to main thread before calling AddCommand because it calls GetService
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                commandService.AddCommand(CreateMenuItem(ApplySettingsCmdId, this.ApplySublimeVSSettings));
            }
        }

        private MenuCommand CreateMenuItem(int cmdId, EventHandler menuItemCallback)
        {
            return new MenuCommand(menuItemCallback, new CommandID(SublimeVSCmdSetGuid, cmdId));
        }

        public void ApplySublimeVSSettings(object sender, EventArgs e)
        {
            ApplySublimeVSSettingsAsync();
        }

        public async Task ApplySublimeVSSettingsAsync()
        {
            // Offer to apply MiniMap (Map Mode Scrollbar)
            await ApplyMiniMapAsync();

            // Offer to apply Sublime shortcut scheme
            await ApplyShortcutsAsync("Sublime Text Shortcuts", SublimeSettingsFileName);
        }

        //-------- MiniMap Settings --------

        public async Task ApplyMiniMapAsync()
        {
            // Confirm with user before applying MiniMap
            if (ConfirmApplyMinimap())
            {
                await ApplyMapModeScrollbarAsync();
            }

        }

        private static bool ConfirmApplyMinimap()
        {
            const string title = "SublimeVS Settings";
            const string message =
                "Activate MiniMap?\n\n" +
                "Apply the following settings:\n" +
                "- Turn on Map Mode Scrollbar (Wide)\n" +
                "\n" +
                "Note: You can modify these settings later at:\n" +
                "- Tools->Options;Text Editor->All Languages->Scroll Bars->Behaviour";
            return MessageBox.Show(message, title, MessageBoxButtons.OKCancel) == DialogResult.OK;
        }

        private async Task ApplyMapModeScrollbarAsync()
        {
            try
            {
                if (await AsyncServiceProvider.GetServiceAsync(typeof(DTE)) is DTE2 dte2)
                {
                    UpdateSetting(dte2, "TextEditor", "AllLanguages", "UseMapMode", true);
                    UpdateSetting(dte2, "TextEditor", "AllLanguages", "OverviewWidth", (short)83);
                }
            }
            catch (Exception e)
            {
                Debug.Print(e.Message);
            }
        }

        private static void UpdateSetting(DTE2 dte2, string category, string page, string settingName, object value)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Example: dte2.Properties["TextEditor", "General"].Item("TrackChanges").Value = true;
            dte2.Properties[category, page].Item(settingName).Value = value;
        }

        //------------ Shortcut Settings --------------

        public async Task ApplyShortcutsAsync(string shortcutSchemeName, string vssettingsFilename)
        {
            //Ask user if they want to apply shortcuts
            if (ConfirmApplyShortcuts(shortcutSchemeName))
            {
                await ImportUserSettingsAsync(vssettingsFilename);
            }
        }

        private bool ConfirmApplyShortcuts(string shortcutSchemeName)
        {
            const string title = "SublimeVS Settings";
            string message =
                $"Apply keyboard shortcuts: {shortcutSchemeName}\n" +
                "\n" +
                "Note: You can modify/reset these settings later at:\n" +
                "- Tools->Options;Environment->Keyboard (Reset)";
            return MessageBox.Show(message, title, MessageBoxButtons.OKCancel) == DialogResult.OK;
        }

        private async Task ImportUserSettingsAsync(string settingsFileName)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            if (await AsyncServiceProvider.GetServiceAsync(typeof(SVsUIShell)) is IVsUIShell shell)
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
