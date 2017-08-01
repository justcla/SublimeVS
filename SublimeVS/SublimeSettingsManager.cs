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

        private readonly Package package;

        private IServiceProvider ServiceProvider
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

        public static void Initialize(Package package)
        {
            Instance = new SublimeSettingsManager(package);
        }

        private SublimeSettingsManager(Package package)
        {
            // Register this command with the Global Command Service
            this.package = package ?? throw new ArgumentNullException("package");

            if (ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                commandService.AddCommand(CreateMenuItem(ApplySettingsCmdId, this.ApplySublimeVSSettings));
            }
        }

        private MenuCommand CreateMenuItem(int cmdId, EventHandler menuItemCallback)
        {
            return new MenuCommand(menuItemCallback, new CommandID(SublimeVSCmdSetGuid, cmdId));
        }

        public void ApplySublimeVSSettings(object sender, EventArgs e)
        {
            ApplySublimeVSSettings();
        }

        public void ApplySublimeVSSettings()
        {
            // Offer to apply Sublime shortcut scheme
            ApplyShortcuts("Sublime Text Shortcuts", SublimeSettingsFileName);

            // Offer to apply MiniMap (Map Mode Scrollbar)
            ApplyMiniMap();
        }

        //-------- MiniMap Settings --------

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
            const string title = "Activate MiniMap";
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

        //------------ Shortcut Settings --------------

        public void ApplyShortcuts(string shortcutSchemeName, string vssettingsFilename)
        {

            //Ask user if they want to apply shortcuts
            if (ConfirmApplyShortcuts(shortcutSchemeName))
            {
                ImportUserSettings(vssettingsFilename);
            }
        }

        private bool ConfirmApplyShortcuts(string shortcutSchemeName)
        {
            const string title = "Apply Shortcuts";
            string message =
                $"Apply keyboard shortcuts: {shortcutSchemeName}\n" +
                "\n" +
                "Note: You can modify/reset these settings later in\n" +
                "- Tools->Options;Environment->Keyboard";
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

    }
}
