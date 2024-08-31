using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointApi
{
    internal class TitleManipulation
    {
        /// <summary>
        /// This Method will Edit the Slide Title as required from task 1
        /// </summary>
        /// <param name="slidePart"></param>
        /// <param name="presnetationDocument"></param>
        /// <param name="newTitle"></param>
        /// <returns>Boolean, true for success and false for failure</returns>
        internal static bool EditSlideTilte(SlidePart slidePart, string newTitle)
        {
            // Get all shapes
            var shapes = slidePart.Slide.Descendants<Shape>();
            foreach (var shape in shapes)
            {
                // Checking if this shape is a title shape
                if (IsTitleShape(shape))
                {
                    // Now we have the title shape
                    var textBody = shape.TextBody;
                    if (textBody != null)
                    {
                        string oldTitle = textBody.InnerText;

                        // Change the title to new title by calling helper method then save the
                        if (IsChangedTitle(textBody, newTitle))
                        {
                            // Align the text in the title to center
                            ChangeAlignmentToCenter(textBody);

                            // Change the font to Beirut
                            ChangeTextFont(textBody, "Beirut");

                            // Since the change was successfull then return a pleasing result
                            return true;
                        }

                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Helper Method to check id the Shape is a title Shape
        /// </summary>
        /// <param name="shape"></param>
        /// <returns>true if shape is title, <br>false otherwise</br></returns>
        internal static bool IsTitleShape(Shape shape)
        {
            var placeholder = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();

            if (placeholder != null && placeholder.Type != null && placeholder.Type.Value == PlaceholderValues.Title)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Helper Method to change the tilte and indicate if the change was successful
        /// </summary>
        /// <param name="textBody"></param>
        /// <param name="newTitle"></param>
        /// <returns>True if changed, <br></br>false otherwise</returns>
        private static bool IsChangedTitle(TextBody textBody, string newTitle)
        {
            var textElement = textBody.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault();

            if (textElement is not null)
            {
                textElement.Text = newTitle;

                return true;
            }

            return false;
        }

        /// <summary>
        /// This Method will change the alignment of the text
        /// </summary>
        /// <param name="textBody"></param>
        /// <param name="newTitle"></param>
        private static void ChangeAlignmentToCenter(TextBody textBody)
        {
            var paragraph = textBody.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();

            if (paragraph is not null)
            {
                // Get Patagraph Properties to access the alignment
                var paragraphProperties = paragraph.ParagraphProperties;

                // Check is Paragraph Properties is null, to create a new one
                if (paragraphProperties is null)
                {
                    paragraphProperties = new DocumentFormat.OpenXml.Drawing.ParagraphProperties();

                    // Insert it into the paragraph elemnt at position 0, since it should be defined as first object
                    paragraph.InsertAt(paragraphProperties, 0);
                }

                // Change the alignemnt to center
                paragraphProperties.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center;
            }
        }

        /// <summary>
        /// This Method changes the Font of a given textBody
        /// </summary>
        /// <param name="textBody"></param>
        internal static void ChangeTextFont(TextBody textBody, string newFont)
        {
            // Get the Runs of the tesxtBody to access RunProperties
            var runs = textBody.Descendants<DocumentFormat.OpenXml.Drawing.Run>();

            // Loop over runs to chnage the font of them
            foreach (var run in runs)
            {
                // Get the run properties or create one if its null
                var runProperties = run.RunProperties;

                // Create a new Property and Append it to the run structure
                if (runProperties is null)
                {
                    runProperties = new DocumentFormat.OpenXml.Drawing.RunProperties();
                    run.AppendChild(runProperties);
                }
                // Get the font preoprty from the RunProperties
                var latinFont = runProperties.GetFirstChild<DocumentFormat.OpenXml.Drawing.LatinFont>();

                // If its null then create one and append it to the RunProperty, and set the font to the given font
                if (latinFont is null)
                {
                    latinFont = new DocumentFormat.OpenXml.Drawing.LatinFont() { Typeface = newFont };
                    runProperties.AppendChild(latinFont);
                }
                else
                {
                    latinFont.Typeface = newFont;
                }

            }
        }
    }
}
