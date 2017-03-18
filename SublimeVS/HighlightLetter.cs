using System;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;

namespace SublimeVS
{
    /// <summary>
    /// HighlightLetter places red boxes behind all the "a"s in the editor window
    /// </summary>
    internal sealed class HighlightLetter
    {
        internal static bool isActive = false;
        internal static bool isWaitingToJump = false;

        /// <summary>
        /// The layer of the adornment.
        /// </summary>
        private readonly IAdornmentLayer layer;

        /// <summary>
        /// Text view where the adornment is created.
        /// </summary>
        private readonly IWpfTextView view;

        /// <summary>
        /// Adornment brush.
        /// </summary>
        private readonly Brush brush;

        /// <summary>
        /// Adornment pen.
        /// </summary>
        private readonly Pen pen;

        /// <summary>
        /// Initializes a new instance of the <see cref="HighlightLetter"/> class.
        /// </summary>
        /// <param name="view">Text view to create the adornment for</param>
        public HighlightLetter(IWpfTextView view)
        {
            if (view == null)
            {
                throw new ArgumentNullException("view");
            }

            this.layer = view.GetAdornmentLayer("HighlightLetter");

            this.view = view;

            // Create the pen and brush to color the box behind the a's
            this.brush = new SolidColorBrush(Color.FromArgb(0x20, 0x00, 0x00, 0xff));
            this.brush.Freeze();

            var penBrush = new SolidColorBrush(Colors.Red);
            penBrush.Freeze();
            this.pen = new Pen(penBrush, 0.5);
            this.pen.Freeze();
        }

        /// <summary>
        /// Handles whenever the text displayed in the view changes by adding the adornment to any reformatted lines
        /// </summary>
        /// <remarks><para>This event is raised whenever the rendered text displayed in the <see cref="ITextView"/> changes.</para>
        /// <para>It is raised whenever the view does a layout (which happens when DisplayTextLineContainingBufferPosition is called or in response to text or classification changes).</para>
        /// <para>It is also raised whenever the view scrolls horizontally or when its size changes.</para>
        /// </remarks>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        internal void TurnOffAndCancel(object sender, TextViewLayoutChangedEventArgs e)
        {
            DeactivateFeature();
        }

        internal void DeactivateFeature()
        {
            if (isActive)
            {
                isActive = false;
                isWaitingToJump = false;
                ClearVisuals();
            }
        }

        internal void ActivateFeature()
        {
            isActive = true;
            isWaitingToJump = false;
            //this.CreateVisuals();
            // If the layout changes (ie. scroll) then cancel the feature
            //this.view.LayoutChanged += this.TurnOffAndCancel;
        }

        internal void HighlightLetters(char letterToHighlight)
        {
            CreateVisuals(letterToHighlight);
            isWaitingToJump = true;
        }

        internal void JumpToChosenPosition(char typedChar)
        {
            if (isActive && isWaitingToJump)
            {
                int newPosition = (int)(typedChar - 'A'); // TODO: Get this from the typedChar
                // Jump to the correct position
                this.view.Caret.MoveTo(new SnapshotPoint(this.view.TextSnapshot, newPosition));
                DeactivateFeature();
            }
        }

        /// <summary>
        /// Adds the scarlet box behind the 'a' characters within the given line
        /// </summary>
        /// <param name="line">Line to add the adornments</param>
        private void CreateVisuals(char letterToHighlight)
        {
            IWpfTextViewLineCollection textViewLines = this.view.TextViewLines;

            foreach (ITextViewLine line in this.view.TextViewLines)
            {
                // Loop through each character, and place a box around any 'a'
                for (int charIndex = line.Start; charIndex < line.End; charIndex++)
                {
                    if (this.view.TextSnapshot[charIndex] == letterToHighlight)
                    {
                        SnapshotSpan span = new SnapshotSpan(this.view.TextSnapshot, Span.FromBounds(charIndex, charIndex + 1));
                        Geometry geometry = textViewLines.GetMarkerGeometry(span);
                        if (geometry != null)
                        {
                            var drawing = new GeometryDrawing(this.brush, this.pen, geometry);
                            drawing.Freeze();

                            var drawingImage = new DrawingImage(drawing);
                            drawingImage.Freeze();

                            var image = new Image
                            {
                                Source = drawingImage,
                            };

                            // Align the image with the top of the bounds of the text geometry
                            Canvas.SetLeft(image, geometry.Bounds.Left);
                            Canvas.SetTop(image, geometry.Bounds.Top);

                            this.layer.AddAdornment(AdornmentPositioningBehavior.TextRelative, span, null, image, null);
                        }
                    }
                }
            }
        }

        private void ClearVisuals()
        {
            this.layer.RemoveAllAdornments();
        }

    }
}
