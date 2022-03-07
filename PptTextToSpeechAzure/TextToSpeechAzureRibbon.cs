using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptTextToSpeechAzure
{
    public partial class TextToSpeechAzureRibbon
    {
        private void TextToSpeechAzureRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        /// <summary>
        /// Seek for for notes in current PowerPoint slide.
        /// </summary>
        /// <returns>String with notes.</returns>
        static private string GetNotesFromCurrentSlide()
        {
            string slideNotes = string.Empty;
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            if (null != slide)
            {
                if (slide.HasNotesPage == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.SlideRange notesPages = slide.NotesPage;
                    foreach (PowerPoint.Shape shape in notesPages.Shapes)
                    {
                        if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            if (shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                            {
                                slideNotes = shape.TextFrame.TextRange.Text;
                                break;
                            }
                        }
                    }
                }
            }

            return slideNotes;
        }

        /// <summary>
        /// Seek for textbox with Title "Caption" in current slide.
        /// </summary>
        /// <returns>Powerpoint Shape object reference.</returns>
        static private PowerPoint.Shape GetCaptionTextboxFromCurrentSlide()
        {
            PowerPoint.Shape captionTextbox = null;
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            if (null != slide)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        string title = shape.Title;
                        if (string.Compare(title, "Caption", StringComparison.Ordinal) == 0)
                        {
                            captionTextbox = shape;
                            break;
                        }
                    }
                }
            }

            return captionTextbox;
        }

        /// <summary>
        /// Add or replace audio object in current slide with title "PptTextToSpeechAzure".
        /// </summary>
        /// <param name="text">Text to become speech.</param>
        /// <param name="voice">Voice name.</param>
        static private void SetAudioForCurrentSlide(string text, string voice)
        {
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            if (null != slide)
            {
                // Seek for and remove previous audio
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    string title = shape.Title;
                    if (string.Compare(title, "SynthSpeech", StringComparison.Ordinal) == 0)
                    {
                        shape.Delete();
                        break;
                    }
                }
            }

            string speechFilename = SpeechSynthesizerHelper.GetSpeech(text, voice);
            PowerPoint.Shape audioShape = slide.Shapes.AddMediaObject2(speechFilename);
            audioShape.Title = "SynthSpeech";
            audioShape.AnimationSettings.PlaySettings.HideWhileNotPlaying = Office.MsoTriState.msoTrue;
            audioShape.AnimationSettings.PlaySettings.PlayOnEntry = Office.MsoTriState.msoTrue;
        }

        private void UpdateCurrentSlide()
        {
            string slideNotes = GetNotesFromCurrentSlide();
            if (!string.IsNullOrEmpty(slideNotes))
            {
                string[] lines = slideNotes.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                // Format is two lines
                // speech
                // voice
                string speech = lines[0];
                string voice = string.Empty;
                if (lines.Length > 1)
                {
                    voice = lines[1];
                }

                PowerPoint.Shape captionTextbox = GetCaptionTextboxFromCurrentSlide();
                if (null != captionTextbox)
                {
                    captionTextbox.TextFrame.TextRange.Text = speech;
                }
                SetAudioForCurrentSlide(speech, voice);
            }
        }

        private void buttonSlideNoteToSpeech_Click(object sender, RibbonControlEventArgs e)
        {
            UpdateCurrentSlide();
        }

        private void buttonUpdateAllSlides_Click(object sender, RibbonControlEventArgs e)
        {
            string slideNotes = string.Empty;
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            if (null != slide)
            {
                if (slide.HasNotesPage == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.SlideRange notesPages = slide.NotesPage;
                    foreach (PowerPoint.Shape shape in notesPages.Shapes)
                    {
                        if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            if (shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                            {
                                slideNotes = shape.TextFrame.TextRange.Text;
                                break;
                            }
                        }
                    }
                }
            }

            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            foreach (PowerPoint.Slide presentationSlide in presentation.Slides)
            {
                presentationSlide.Select();
                UpdateCurrentSlide();
            }
        }
    }
}
