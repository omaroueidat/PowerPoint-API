using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace PowerPointApi
{
    internal class BulletPointManipulation
    {
        /// <summary>
        /// This is the main method that will change the bullet points according to the tasks
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="newFont"></param>
        internal static void ManipulateBulletPointTextBoxes(SlidePart slide, string newFont)
        {
            var bulltetTextBoxes = GetTextBoxes(slide);

            // Get the coordinates of the first Bullet TextBox
            var firstTextBoxCoordiantes = bulltetTextBoxes[0].ShapeProperties.Transform2D;

            foreach(var textBox in bulltetTextBoxes)
            {
                // Chnage the Font
                TitleManipulation.ChangeTextFont(textBox.TextBody, newFont);

                // Remove Bold
                RemoveBold(textBox.TextBody);

                // Remove the Underline 
                RemoveUnderline(textBox.TextBody);

                // Align them on the same level
                ShapeManipulation.AlignAndResizeShapes(textBox, firstTextBoxCoordiantes.Extents.Cx, firstTextBoxCoordiantes.Extents.Cy, firstTextBoxCoordiantes.Offset.Y, 0);

                // Change Bullet Points to a dot
                ChangeBulletPoints(textBox);
            }
        }

        private static List<Shape> GetTextBoxes(SlidePart slide)
        {
            var shapes = slide.Slide.Descendants<Shape>().ToList();

            var bulletTextBoxes = new List<Shape>();

            if (shapes is null)
            {
                return null;
            }

            foreach(var shape in shapes)
            {
                if (TitleManipulation.IsTitleShape(shape))
                {
                    continue;
                }

                if (CheckTextBox(shape))
                {
                    bulletTextBoxes.Add(shape);
                }
            }

            return bulletTextBoxes;
        }

        private static void RemoveBold(TextBody textBody)
        {
            var runs = textBody.Descendants<Drawing.Run>();

            foreach(var run in runs)
            {
                // Get the run properties or create one if its null
                var runProperties = run.RunProperties;

                // Create a new Property and Append it to the run structure
                if (runProperties is null)
                {
                    runProperties = new Drawing.RunProperties();
                    run.AppendChild(runProperties);
                }

                runProperties.Bold = false;
            }
        }

        private static void RemoveUnderline(TextBody textBody)
        {
            var runs = textBody.Descendants<Drawing.Run>();

            foreach (var run in runs)
            {
                // Get the run properties or create one if its null
                var runProperties = run.RunProperties;

                // Create a new Property and Append it to the run structure
                if (runProperties is null)
                {
                    runProperties = new Drawing.RunProperties();
                    run.AppendChild(runProperties);
                }

                runProperties.Underline = null;
            }
        }

        private static bool CheckTextBox(Shape shape)
        {
            var presetGeometry = shape.ShapeProperties.GetFirstChild<Drawing.PresetGeometry>();

            // Check if the shape has a TextBody and if its geometry is a rectangle
                return shape.TextBody != null &&
                       shape.ShapeProperties != null &&
                       shape.ShapeProperties.GetFirstChild<Drawing.PresetGeometry>() != null &&
                       presetGeometry.Preset == Drawing.ShapeTypeValues.Rectangle;
        }

        private static void ChangeBulletPoints(Shape textBox)
        {
            // Ensure the shape has a TextBody
            if (textBox.TextBody == null)
            {
                return;
            }

            // Iterate through all paragraphs within the TextBody
            foreach (var paragraph in textBox.TextBody.Descendants<Drawing.Paragraph>())
            {
                // Get or create ParagraphProperties
                var paragraphProperties = paragraph.GetFirstChild<Drawing.ParagraphProperties>();
                if (paragraphProperties == null)
                {
                    paragraphProperties = new Drawing.ParagraphProperties();
                    paragraph.PrependChild(paragraphProperties);
                }

                // Check is it have numbered vullets
                var isNumbered = paragraphProperties.Descendants<Drawing.AutoNumberedBullet>().FirstOrDefault();

                if (isNumbered is not null)
                {
                    // Remove numbering and add Character Bullet
                    paragraphProperties.RemoveChild(isNumbered);
                    paragraphProperties.AppendChild(new Drawing.CharacterBullet() { Char = "•" });
                }
                else
                {
                    // Set the Bullet to a point
                    var characterBullet = paragraphProperties.Descendants<Drawing.CharacterBullet>().FirstOrDefault();
                    characterBullet.Char = "•";
                }

                
            }
        }
    }
}
