using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio;

namespace VSSettingsManager
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(VSSettingsManagerPackage.PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class VSSettingsManagerPackage : Package
    {
        /// <summary>
        /// VSSettingsManagerPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "2145fd5c-c814-4772-b19d-b840113afede";

        public VSSettingsManagerPackage()
        {
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            // Adds commands handlers for the VS Settings operations (Apply, Backup, Restore, Reset)
            VSSettingsManager.Initialize(this);
            base.Initialize();
        }

    }
}
