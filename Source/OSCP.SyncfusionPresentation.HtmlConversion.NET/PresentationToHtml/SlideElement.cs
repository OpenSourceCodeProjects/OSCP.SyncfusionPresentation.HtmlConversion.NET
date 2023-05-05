using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class SlideElement : HtmlElement
    {
        internal const string ELEMENT_NAME = "div";

        internal HtmlElement Parent { get; set; }
        internal ISlide Slide { get; set; }

        public SlideElement(XmlNode node) : base (node) 
        {
            this.AddClass(HtmlDocument.Settings.CssClass.Slide);
        }

        internal SlideElement Load(ISlide slide)
        {
            this.Slide = slide;

            this.Css("position", "relative")
                .Css("height", $"{slide.SlideSize.Height}px")
                .Css("width", $"{slide.SlideSize.Width}px");

            // If the slide has a background, then apply it to the slide element.
            this.ApplyBackground(slide);

            // Loop over all the shapes in the slide.
            foreach (IShape shape in slide.Shapes)
            {
                if (shape.SlideItemType == SlideItemType.Placeholder)
                {
                    // Create a placeholder element.
                    PlaceholderElement placeholderElement = this.AppendElement<PlaceholderElement>(AutoShapeElement.ELEMENT_NAME);
                    placeholderElement.Parent = this;

                    // Load the placeholder element from the Syncfusion shape.
                    placeholderElement.Load(shape);
                }
                else if (shape.SlideItemType == SlideItemType.Table)
                {
                    // Create a table element.
                    TableElement tableElement = this.AppendElement<TableElement>(TableElement.ELEMENT_NAME);
                    tableElement.Parent = this;

                    // Load the table element from the Syncfusion table.
                    tableElement.Load(shape as ITable);
                }
                else if (shape.SlideItemType == SlideItemType.Picture)
                {
                    // Create a picture element.
                    PictureElement pictureElement = this.AppendElement<PictureElement>(PictureElement.ELEMENT_NAME);
                    pictureElement.Parent = this;

                    // Load the picture element from the Syncfusion picture.
                    pictureElement.Load(shape as IPicture);
                }
                else if (shape.SlideItemType == SlideItemType.AutoShape)
                {
                    // Create an autoshape element.
                    AutoShapeElement autoshapeElement = this.AppendElement<AutoShapeElement>(AutoShapeElement.ELEMENT_NAME);
                    autoshapeElement.Parent = this;

                    // Load the autoshape element from the Syncfusion shape.
                    autoshapeElement.Load(shape);
                }
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
                            string imageAsBase64 = System.Convert.ToBase64String(slide.Background.Fill.PictureFill.ImageBytes);

                            if (HtmlDocument.Settings.ImageData.Location == ImageDataLocation.Embedded)
                            {
                                this.Css("background-image", $"url(data:image/{imageType};base64,{imageAsBase64})")
                                    .Css("background-size", $"{slide.SlideSize.Width}px {slide.SlideSize.Height}px");
                            }
                            else if (HtmlDocument.Settings.ImageData.Location == ImageDataLocation.Base64)
                            {
                                string imageId = Guid.NewGuid().ToString();
                                this.Attr("data-image-id", imageId);
                                HtmlDocument.Settings.ImageData.Base64Images.Add(new Base64Image
                                {
                                    Id = imageId,
                                    ElementName = SlideElement.ELEMENT_NAME,
                                    ElementCssClass = HtmlDocument.Settings.CssClass.Slide,
                                    ImageType = imageType,
                                    Data = imageAsBase64,
                                    Width = slide.SlideSize.Width,
                                    Height = slide.SlideSize.Height
                                });
                            }
                        }
                    }
                }
                else if (slide.Background.Fill.FillType == FillType.Solid) 
                {
                    //IColor color = slide.Background.Fill.SolidFill.Color;
                    //this.Css("background-color", $"rgba({color.R},{color.G},{color.B},{(color.A == 0 ? "1" : "0." + (100F - ((color.A / 255F) * 100F)))})");
                    this.Css("background-color", slide.Background.Fill.SolidFill.Color.SystemColor.ToCss());
                }
            }
        }
    }
}
