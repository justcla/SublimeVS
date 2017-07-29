using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SublimeVS
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(SublimeVSPackage.PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class SublimeVSPackage : Package
    {
        public const string PackageGuidString = "b303a85f-1765-435e-8b09-b600853edbef";
        private const string SID_SVsSettingsPersistenceManager = "9B164E40-C3A2-4363-9BC5-EB4039DEF653";

        private const string settingsFileName = @"Shortcuts\SublimeShortcuts.vssettings";

        public static ISettingsManager SettingsManager { get; private set; }

        public SublimeVSPackage()
        {
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            SublimeVSPackage.SettingsManager = (ISettingsManager)GetGlobalService(typeof(SVsSettingsPersistenceManager));
            base.Initialize();

            // Check if we need to do first-time setup
            const string firstTimeRunSettingName = "SublimeSettingsPrompted";
            if ((SublimeVSPackage.SettingsManager.TryGetValue(firstTimeRunSettingName, out bool value) != GetValueResult.Success) || !value)
            {
                SublimeVSPackage.SettingsManager.SetValueAsync(firstTimeRunSettingName, true, isMachineLocal: true);
                ApplyFirstTimeSettings();
            }

        }

        private void ApplyFirstTimeSettings()
        {

            //Ask user if they want to apply Sublime settings
            const string title = "Extensions for Accessibility in Visual Studio";
            const string message =
                "Congratulations! Sublime VS Extension has been installed.\n" +
                "\n" +
                "Apply the following settings:\n" +
                "- Turn on Map Mode Scrollbar (Wide)\n" +
                "- Assign shortcut Ctrl+P to GoToFile\n" +
                "- Assign shortcut Ctrl+Shift+P to Command Window\n" +
                "\n" +
                "Note: You can modify these settings later in Tools->Options";

            // Show a message box to prove we were here
            int result = VsShellUtilities.ShowMessageBox(
                this,
                message,
                title,
                OLEMSGICON.OLEMSGICON_WARNING,
                OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            //if (MessageBox.Show(message, title, MessageBoxButtons.OKCancel) == DialogResult.OK)
            if (result == 1)
            {
                ApplySettings();
            }
        }

        private void ApplySettings()
        {
            try
            {
                var provider = (System.IServiceProvider)this;
                var dte2 = (DTE2)(provider.GetService(typeof(DTE)));

                ActivateMapModeScrollbar(dte2);
                //ApplyShortcuts(dte2);
                ImportSublimeShortcuts();
                // TODO: Output to the status bar - "SublimeVS settings applied"
                SublimeVSPackage.SettingsManager.SetValueAsync("SublimeSettingsApplied", true, isMachineLocal: true);
            }
            catch (Exception e)
            {
                Debug.Print(e.Message);
            }
        }

        private void ImportSublimeShortcuts()
        {
            if (this.GetService(typeof(SVsUIShell)) is IVsUIShell shell)
            {
                // import the settings file into Visual Studio
                var asmDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var settingsFilePath = Path.Combine(asmDirectory, settingsFileName);
                var group = VSConstants.CMDSETID.StandardCommandSet2K_guid;

                object arguments = string.Format(CultureInfo.InvariantCulture, "-import:\"{0}\"", settingsFilePath);
                shell.PostExecCommand(ref group, (uint)VSConstants.VSStd2KCmdID.ManageUserSettings, 0, ref arguments);
            }
        }

        private void ActivateMapModeScrollbar(DTE2 dte2)
        {
            UpdateSetting(dte2, "TextEditor", "AllLanguages", "UseMapMode", true);
            UpdateSetting(dte2, "TextEditor", "AllLanguages", "OverviewWidth", (short)83);
        }

        private static void UpdateSetting(DTE2 dte2, string category, string page, string settingName, object value)
        {
            // Example: dte2.Properties["TextEditor", "General"].Item("TrackChanges").Value = true;
            dte2.Properties[category, page].Item(settingName).Value = value;
        }

        internal static void ApplyShortcuts(DTE2 dte2)
        {
            // Apply the shortcuts
            Commands cmds = dte2.Commands;
            AddKeyBinding(cmds, "Edit.GoToFile", "Global", "Ctrl+P");
            ReplaceKeyBinding(cmds, "Edit.GoToFile", "Text Editor", "Ctrl+P");  // Override any Text Editor binding for this Global shortcut
            AddKeyBinding(cmds, "View.CommandWindow", "Global", "Ctrl+Shift+P");
            ReplaceKeyBinding(cmds, "View.CommandWindow", "Text Editor", "Ctrl+Shift+P");  // Override any Text Editor binding for this Global shortcut
        }

        private static void AddKeyBinding(Commands cmds, string vsCommandName, string scope, string keyBinding)
        {
            Command command = cmds.Item(vsCommandName);
            command.Bindings = (object)AppendKeyboardBinding(command, $"{scope}::{keyBinding}");
        }

        private static void ReplaceKeyBinding(Commands cmds, string vsCommandName, string scope, string keyBinding)
        {
            Command command = cmds.Item(vsCommandName);
            // Build new array with just the new binding.
            command.Bindings = new object[] { $"{scope}::{keyBinding}" };
        }

        private static object[] AppendKeyboardBinding(Command command, string keyboardBindingDefn)
        {
            object[] oldBindings = (object[])command.Bindings;

            // Check that keyboard binding is not already there
            for (int i = 0; i < oldBindings.Length; i++)
            {
                if (keyboardBindingDefn.Equals(oldBindings[i]))
                {
                    // Exit early and return the existing bindings array if new keyboard binding is already there
                    return oldBindings;
                }
            }

            // Build new array with all the old bindings, plus the new one.
            object[] newBindings = new object[oldBindings.Length + 1];
            Array.Copy(oldBindings, newBindings, oldBindings.Length);
            newBindings[newBindings.Length - 1] = keyboardBindingDefn;
            return newBindings;
        }

        // A horrible hack but SVsSettingsPersistenceManager isn't public and we need something with the right GUID to get the service.
        [Guid(SID_SVsSettingsPersistenceManager)]
        private class SVsSettingsPersistenceManager
        { }
    }
}
