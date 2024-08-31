using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Net.WebSockets;
using System.Text;


namespace PowerPointApi
{
    class MainProgram
    {
        static void Main(string[] args)
        {
            // Path to input power point file
            string inputFile = @"C:\Users\user\Desktop\WorkShop\.Net Development\Projects\PowerPointApi\PowerPointApi\AuxiTask.pptx";

            // Opening the PowerPoint
            PresentationDocument presentation = PresentationDocument.Open(inputFile, true);
            PresentationPart presentationPart = presentation.PresentationPart;

            // Get the first Slide
            var firstSlide = presentationPart.SlideParts.ElementAt(0);

            // Edit the title name, alignment and font
            TitleManipulation.EditSlideTilte(firstSlide, "Output Slide");
            

            

            // Merge the text and shapes
            ShapeManipulation.MergeTextAndShapes(firstSlide);

            // Edit the Bullet Points
            BulletPointManipulation.ManipulateBulletPointTextBoxes(firstSlide, "Beirut");

            // Save the changes
            firstSlide.Slide.Save();
            presentation.Save();

        }
    }
}