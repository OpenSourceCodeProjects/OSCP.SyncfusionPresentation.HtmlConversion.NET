using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    public class SlideElement : HtmlElement
    {
        internal HtmlElement Parent { get; set; }
        internal ISlide Slide { get; set; }

        public SlideElement(XmlNode node) : base (node) 
        {
            this.AddClass("pptx-slide");
        }

        internal SlideElement Load(ISlide slide)
        {
            this.Slide = slide;

            this.Css("border", "1px solid black")
                .Css("position", "relative")
                .Css("height", $"{slide.SlideSize.Height}px")
                .Css("width", $"{slide.SlideSize.Width}px");

            // If the slide has a background, then apply it to the slide element.
            this.ApplyBackground(slide);

            // Loop over all the shapes in the slide.
            foreach (IShape shape in slide.Shapes)
            {
                // Create a shape element.
                ShapeElement shapeElement = this.AddElement<ShapeElement>("div");
                shapeElement.Parent = this;

                // Load the shape element from the Syncfusion shape.
                shapeElement.Load(shape);
            }

            // Update the underlying XmlNode with the styles and classes.
            this.Update();

            return this;
        }

        /// <summary>
        /// Determine whether the slide has a background and if so, then apply it.
        /// </summary>
        /// <param name="slide">Syncfusion slide.</param>
        private void ApplyBackground(ISlide slide)
        {
            // The slide has a background.
            if (slide.Background.Fill.FillType != FillType.None)
            {
                if (slide.Background.Fill.FillType == FillType.Picture)
                {
                    // Get the picture and add as the slide backgound as a base64 string.
                    using (MemoryStream stream = new MemoryStream(slide.Background.Fill.PictureFill.ImageBytes))
                    {
                        // Figure out the image type.
                        System.Drawing.Image image = System.Drawing.Image.FromStream(stream);
                        string imageType = new System.Drawing.ImageFormatConverter().ConvertToString(image.RawFormat).ToLower();

                        // Got the image type.
                        if (imageType.Length > 0)
                        {
                            // Add the background image as a base64 string.
                            string imageAsBase64 = $"url(data:image/{imageType};base64,{System.Convert.ToBase64String(slide.Background.Fill.PictureFill.ImageBytes)})";
                            this.Css("background-image", imageAsBase64)
                                .Css("background-size", $"{slide.SlideSize.Width}px {slide.SlideSize.Height}px");
                        }
                    }
                }
            }
        }
    }
}
