using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Operations;

namespace SublimeVS
{
    /// <summary>
    /// Establishes an <see cref="IAdornmentLayer"/> to place the adornment on and exports the <see cref="IWpfTextViewCreationListener"/>
    /// that instantiates the adornment on the event of a <see cref="IWpfTextView"/>'s creation
    /// </summary>
    [Export(typeof(IVsTextViewCreationListener))]
    [ContentType("text")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    internal sealed class HighlightLetterTextViewCreationListener : IVsTextViewCreationListener
    {
        // Disable "Field is never assigned to..." and "Field is never used" compiler's warnings. Justification: the field is used by MEF.
#pragma warning disable 649, 169

        [Import]
        private IVsEditorAdaptersFactoryService EditorAdaptersFactoryService { get; set; }

        [Import]
        private IClassifierAggregatorService _aggregatorFactory;

        [Import]
        private SVsServiceProvider _globalServiceProvider;

        [Import(typeof(IEditorOperationsFactoryService))]
        private IEditorOperationsFactoryService _editorOperationsFactory;


#pragma warning restore 649, 169

        #region IWpfTextViewCreationListener

        /// <summary>
        /// Called when a text view having matching roles is created over a text data model having a matching content type.
        /// Instantiates a HighlightLetter manager when the textView is created.
        /// </summary>
        /// <param name="textView">The <see cref="IWpfTextView"/> upon which the adornment should be placed</param>
        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            IWpfTextView textView = EditorAdaptersFactoryService.GetWpfTextView(textViewAdapter);

            // The adornment will listen to any event that changes the layout (text changes, scrolling, etc)
            var highlightLetter = new HighlightLetter(textView);

            CommandFilter commandFilter = new CommandFilter(textView, _aggregatorFactory, _globalServiceProvider, _editorOperationsFactory);
            //CommandFilter commandFilter = new CommandFilter(textView);
            IOleCommandTarget next;
            textViewAdapter.AddCommandFilter(commandFilter, out next);

            commandFilter.Next = next;

        }

        #endregion
    }
}
