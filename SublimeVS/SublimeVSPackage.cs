using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace SublimeVS
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.2.1", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(SublimeVSPackage.PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string, PackageAutoLoadFlags.BackgroundLoad)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class SublimeVSPackage : AsyncPackage
    {
        public const string PackageGuidString = "10faf7a3-f1bb-4836-9e6b-b5f52bd88031";
        private const string SID_SVsSettingsPersistenceManager = "9B164E40-C3A2-4363-9BC5-EB4039DEF653";

        //private const string SublimeSettingsFileName = @"Shortcuts\SublimeShortcuts.vssettings";

        public static ISettingsManager SettingsManager { get; private set; }

        public SublimeVSPackage()
        {
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            // Initialize settings manager (TODO: could be done lazily on get)
            SettingsManager = (ISettingsManager)GetGlobalService(typeof(SVsSettingsPersistenceManager));

            await SublimeSettingsManager.InitializeAsync(this);
            await CheckFirstTimeSetupAsync();
        }

        private async System.Threading.Tasks.Task CheckFirstTimeSetupAsync()
        {
            // Check if we need to do first-time setup
            const string firstTimeRunSettingName = "SublimeVSSetupAck02";
            if ((SettingsManager.TryGetValue(firstTimeRunSettingName, out bool value) != GetValueResult.Success) || !value)
            {
                const string title = "Sublime VS - First Time Setup";
                const string message =
                    "Congratulations! Sublime VS Extension has been installed.\n" +
                    "\n" +
                    "You will now be prompted to apply the following settings:\n" +
                    "- Turn on Map Mode Scrollbar (Wide)\n" +
                    "- Apply Sublime Text keyboard shortcuts\n" +
                    "\n" +
                    "Note: You can modify these settings later in Tools->Options\n" +
                    "\n" +
                    "Click Cancel to Show this dialog again at next startup.";

                if (MessageBox.Show(message, title, MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    // Apply First Time Settings();
                    await SettingsManager.SetValueAsync(firstTimeRunSettingName, true, isMachineLocal: true);

                    await SublimeSettingsManager.Instance.ApplySublimeVSSettingsAsync();
                }
            }
        }

        // A horrible hack but SVsSettingsPersistenceManager isn't public and we need something with the right GUID to get the service.
        [Guid(SID_SVsSettingsPersistenceManager)]
        private class SVsSettingsPersistenceManager
        { }
    }
}
