using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using System;
using System.Runtime.InteropServices;
using OLEConstants = Microsoft.VisualStudio.OLE.Interop.Constants;

namespace SublimeVS
{
    class CommandFilter : IOleCommandTarget
    {

        private readonly IWpfTextView textView;
        private readonly IClassifier classifier;
        private readonly SVsServiceProvider globalServiceProvider;
        private IEditorOperations editorOperations;

        private bool jumpToLetterMode = false;
        private HighlightLetter HighlightLetter;

        public CommandFilter(IWpfTextView textView, IClassifierAggregatorService aggregatorFactory,
            SVsServiceProvider globalServiceProvider, IEditorOperationsFactoryService editorOperationsFactory)
        {
            this.textView = textView;
            classifier = aggregatorFactory.GetClassifier(textView.TextBuffer);
            this.globalServiceProvider = globalServiceProvider;
            editorOperations = editorOperationsFactory.GetEditorOperations(textView);
        }

        public IOleCommandTarget Next { get; internal set; }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            // Command handling
            if (pguidCmdGroup == Constants.SublimeVSGuid)
            {
                // Dispatch to the correct command handler
                switch (nCmdID)
                {
                    case Constants.JumpToLetterCmdId:
                        return HandleJumpToLetter(textView, classifier, GetShellCommandDispatcher(), editorOperations);
                }
            }

            if (HighlightLetter.isActive)
            {
                char typedChar = char.MinValue;

                if (pguidCmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.TYPECHAR)
                {
                    typedChar = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);
                    return HandleSelectLetterToHighlight(textView, typedChar, classifier, GetShellCommandDispatcher(), editorOperations);
                }
            }

            // No commands called. Pass to next command handler.
            if (Next != null)
            {
                return Next.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            }
            return (int)OLEConstants.OLECMDERR_E_UNKNOWNGROUP;
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            // Command handling registration
            if (pguidCmdGroup == Constants.SublimeVSGuid && cCmds == 1)
            {
                switch (prgCmds[0].cmdID)
                {
                    case Constants.JumpToLetterCmdId:
                        prgCmds[0].cmdf |= (uint)OLECMDF.OLECMDF_ENABLED;
                        return VSConstants.S_OK;
                }
            }

            if (Next != null)
            {
                return Next.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
            }
            return (int)OLEConstants.OLECMDERR_E_UNKNOWNGROUP;
        }

        /// <summary>
        /// Get the SUIHostCommandDispatcher from the global service provider.
        /// </summary>
        private IOleCommandTarget GetShellCommandDispatcher()
        {
            return globalServiceProvider.GetService(typeof(SUIHostCommandDispatcher)) as IOleCommandTarget;
        }

        //--- Event Handlers ----//

        private int HandleJumpToLetter(IWpfTextView textView, IClassifier classifier, IOleCommandTarget oleCommandTarget, IEditorOperations editorOperations)
        {
            HighlightLetter highlightLetterController = textView.Properties["HighlightLetterLayer"] as HighlightLetter;
            if (!HighlightLetter.isActive)
            {
                highlightLetterController.ActivateFeature();
            } else
            {
                highlightLetterController.DeactivateFeature();
            }
            return VSConstants.S_OK;
        }

        private int HandleSelectLetterToHighlight(IWpfTextView textView, char typedChar, IClassifier classifier, IOleCommandTarget oleCommandTarget, IEditorOperations editorOperations)
        {
            HighlightLetter highlightLetterController = textView.Properties["HighlightLetterLayer"] as HighlightLetter;
            if (HighlightLetter.isActive)
            {
                highlightLetterController.HighlightLetters(typedChar);
            }
            return VSConstants.S_OK;
        }

    }
}
