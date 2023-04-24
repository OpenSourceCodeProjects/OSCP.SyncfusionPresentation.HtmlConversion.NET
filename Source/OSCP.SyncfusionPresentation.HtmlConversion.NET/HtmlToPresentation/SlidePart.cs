using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    /// <summary>
    /// Convert an HTMLDivElement with the pptx-slide class to a Syncfusion.Presentaton.ISlide object.
    /// </summary>
    internal class SlidePart : PartObject
    {
        /// <summary>
        /// Parent HTML presentation document.
        /// </summary>
        internal PptxDocument Document { set; get; }

        /// <summary>
        /// Syncfusion.Presentaton.ISlide object that is created.
        /// </summary>
        internal ISlide ISlide { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="document">Parent HTML presentation document.</param>
        internal SlidePart(PptxDocument document)
        {
            this.Document = document;
        }

        /// <summary>
        /// Load the slide.
        /// </summary>
        /// <param name="slideNode">XmlNode that represents the HTMLDivElement for the slide.</param>
        internal void Load(XmlNode slideNode)
        {
            this.Node = slideNode;

            // Adds a slide to the PowerPoint presentation.
            this.ISlide = this.Document.IPresentation.Slides.Add(SlideLayoutType.Blank);

            // Apply a background.
            this.ApplyBackground();

            // Get the shapes.
            XmlNodeList shapes = this.Node.SelectNodes("div[contains(@class, 'pptx-shape')]");

            // Loop over all the shapes in the slide.
            foreach (XmlNode shapeNode in shapes)
            {
                // Create and load a shape.
                ShapePart shapePart = new ShapePart(this);
                shapePart.Load(shapeNode);
            }

            // Get the tables.
            XmlNodeList tables = this.Node.SelectNodes("table[contains(@class, 'pptx-table')]");

            // Loop over all the tables in the slode.
            foreach (XmlNode tableNode in tables)
            {
                // Create and load a table.
                TablePart tablePart = new TablePart(this);
                tablePart.Load(tableNode);
            }
        }

        /// <summary>
        /// Determine whether the slide has a background and if so, then apply it.
        /// </summary>
        private void ApplyBackground()
        {
            string styleValue = string.Empty;

            // Retrieve the background instance.
            IBackground background = this.ISlide.Background;

            // Solid background color.
            if ((styleValue = this.Css("background-color")).Length > 0) 
            {
                ColorObject colorObject = ColorExtensions.FromCss(styleValue);

                // Not 100% transparent.
                if (colorObject.A != 255)
                {
                    background.Fill.FillType = FillType.Solid;
                    background.Fill.SolidFill.Color = colorObject;
                }
            }
            // Background image.
            else if ((styleValue = this.Css("background-image")).Length > 0)
            {
                // Get the base 64 string with the image.
                Match imageMatch = Regex.Match(styleValue, @"url\(data:image/[^;]*;base64,([^\)]*)\)");

                // A base 64 string for the image was found.
                if (imageMatch.Success == true) 
                {
                    // Get the base64 string and apply the background image.
                    string base64ImageString = imageMatch.Groups[1].Value;
                    byte[] imageBytes = Convert.FromBase64String(base64ImageString).ToArray();
                    background.Fill.FillType = FillType.Picture;
                    background.Fill.PictureFill.ImageBytes = imageBytes;
                }
            }
        }
    }
}
