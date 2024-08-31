using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace PowerPointApi
{
    internal class ShapeManipulation
    {
        /// <summary>
        /// Main Method that will be called to do the task of changing manipulating the shapes
        /// </summary>
        /// <param name="slidePart"></param>
        internal static void MergeTextAndShapes(SlidePart slidePart)
        {
            // Get All Shapes
            var shapes = slidePart.Slide.Descendants<Shape>();

            // Collection of shapes to be removed
            var shapesToRemove = new List<Shape>();

            foreach (var textBoxShape in shapes)
            {
                // Skip the title
                if (TitleManipulation.IsTitleShape(textBoxShape))
                {
                    continue;
                }

                // Skip the Shapes without textBody to insure that only textBoxes are selected
                var textBody = textBoxShape.TextBody;
                if (textBody is null || string.IsNullOrEmpty(textBody.InnerText))
                {
                    continue;
                }

                // Find the Shape below the textBox, and if there is no shape, then skip because it is not what we want to manipualte
                var shapeBelowTextBox = shapes.FirstOrDefault(shape =>
                {
                    // Check if we are comparing a shape to itself, which must not happen
                    if (shape == textBoxShape)
                    {
                        return false;
                    }

                    // Check is the shape is containing the textBoxShape
                    return IsShapeContainAnother(shape, textBoxShape);
                });

                if (shapeBelowTextBox is not null)
                {
                    ChangeTextInShape(shapeBelowTextBox, textBody.InnerText);

                    // Align Shapes
                    AlignAndResizeShapes(shapeBelowTextBox, 3000000, 1500000, 1712339, 500000);


                    // Add the textBox to the collection to remove it
                    shapesToRemove.Add(textBoxShape);


                }
            }

            // Remove all the shapes from the collection of shapes
            foreach (var shape in shapesToRemove)
            {
                shape.Remove();
            }
        }

        /// <summary>
        /// This Methos Checks if a Shape (Parent) Contains Antoher Shape (Child), by checking the coordinates of the shapes 
        /// </summary>
        /// <param name="parentShape"></param>
        /// <param name="childShape"></param>
        /// <returns>True if Parent Contains Child, false otherwise</returns>
        private static bool IsShapeContainAnother(Shape parentShape, Shape childShape)
        {
            // Get the ShapeProperties elements
            var parentShapeProperties = parentShape.ShapeProperties;
            var childShapeProperties = childShape.ShapeProperties;

            // Check if anything is null
            if (parentShapeProperties is null || childShapeProperties is null)
            {
                return false;
            }

            // Get parent shape coordinates from ShapeProperties
            var parentTransformShape = parentShapeProperties.Transform2D;
            var childTransformShape = childShapeProperties.Transform2D;

            if (parentTransformShape is null || childTransformShape is null)
            {
                return false;
            }

            // Get parent shape coordinates. Type is Int64
            var parentStartX = parentTransformShape.Offset.X.Value;
            var parentStartY = parentTransformShape.Offset.Y.Value;
            var parentEndX = parentStartX + parentTransformShape.Extents.Cx.Value;
            var parentEndY = parentStartY + parentTransformShape.Extents.Cy.Value;

            // Get child shape coordinates
            var childStartX = childTransformShape.Offset.X.Value;
            var childStartY = childTransformShape.Offset.Y.Value;
            var childEndX = childStartX + childTransformShape.Extents.Cx.Value;
            var childEndY = childStartY + childTransformShape.Extents.Cy.Value;

            return (parentStartX <= childStartX &&
                    parentStartY <= childStartY &&
                    parentEndX >= childEndX &&
                    parentEndY >= childEndY
                   );
        }

        /// <summary>
        /// Helper Method to add the text to the shapes
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="newText"></param>
        private static void ChangeTextInShape(Shape shape, string newText)
        {
            // Ensure the shape has a TextBody
            if (shape.TextBody == null)
            {
                shape.TextBody = new TextBody(new DocumentFormat.OpenXml.Drawing.BodyProperties(), new DocumentFormat.OpenXml.Drawing.ListStyle());

                // Create a new paragraph
                Paragraph paragraph = new Paragraph();

                // Create a new run
                Run run = new Run();

                // Create the text element with the new text
                DocumentFormat.OpenXml.Drawing.Text textElement = new DocumentFormat.OpenXml.Drawing.Text(newText);

                // Append the text element to the run
                run.Append(textElement);

                // Append the run to the paragraph
                paragraph.Append(run);

                // Append the paragraph to the TextBody
                shape.TextBody.Append(paragraph);

            }
            else
            {

                // Get the paragraph
                var paragraph = shape.TextBody.LastChild;


                // Create the text element with the new text
                DocumentFormat.OpenXml.Drawing.Text textElement = new DocumentFormat.OpenXml.Drawing.Text(newText);


                // Append the run to the paragraph
                paragraph.AppendChild(textElement);


            }



        }

        /// <summary>
        /// Method to align all the shapes on the same Y-axis Level and with equal size
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="alignment"></param>
        /// <param name="spacing"></param>
        internal static void AlignAndResizeShapes(Shape shape, long width, long height, long alignment, long spacing)
        {
            var shapeProperties = shape.ShapeProperties;

            var shapeTranform = shapeProperties.Transform2D;

            // Set the shape size to given size
            shapeTranform.Extents.Cx = width;
            shapeTranform.Extents.Cy = height;

            // Align the shapes all on the same level
            shapeTranform.Offset.Y = alignment;
        }
    }
}
