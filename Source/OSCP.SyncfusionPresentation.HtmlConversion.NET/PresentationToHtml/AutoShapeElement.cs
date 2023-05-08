using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class AutoShapeElement : SlideItemElement
    {
        internal const string ELEMENT_NAME = "div";

        internal IShape Shape { get; set; }
        protected string CssClassSettings { get; set; }
        protected Dictionary<string, string> ElementAttributes { get; set; }

        public AutoShapeElement(XmlNode node) : base(node)
        {
            this.CssClassSettings = HtmlDocument.Settings.CssClass.AutoShape;

            this.ElementAttributes = HtmlDocument.Settings.ElementAttributes.AutoShape;
        }

        internal AutoShapeElement Load(IShape shape)
        {
            this.AddClass(this.CssClassSettings);

            this.Shape = shape;

            // Position the element on the slide.
            this.PositionElement(shape);

            // The shape is an image.
            if (shape.GetType().Name == "Picture")
            {
                // Add the css class for a picture.
                this.AddClass(HtmlDocument.Settings.CssClass.Picture);

                // Cast the shape as a picture to get the image information.
                IPicture picture = shape as IPicture;

                // Get the image type.
                string imageType = picture.ImageFormat.ToString().ToLower();

                // Get the image as a base64 string.
                string imageAsBase64 = System.Convert.ToBase64String(picture.ImageData);

                if (HtmlDocument.Settings.ImageData.Location == ImageDataLocation.Embedded)
                {
                    // Set the background of this shape with the image.
                    this.Css("display", "initial")
                        .Css("background-image", $"url(data:image/{imageType};base64,{imageAsBase64})")
                        .Css("background-size", $"{picture.Width}px {picture.Height}px")
                        .Css("background-repeat", "no-repeat");
                }
                else if (HtmlDocument.Settings.ImageData.Location == ImageDataLocation.Base64)
                {
                    this.Css("display", "initial");

                    string imageId = Guid.NewGuid().ToString();
                    this.Attr("data-image-id", imageId);

                    HtmlDocument.Settings.ImageData.Base64Images.Add(new Base64Image
                    {
                        Id = imageId,
                        ElementName = AutoShapeElement.ELEMENT_NAME,
                        ElementCssClass = this.CssClassSettings,
                        ImageType = imageType,
                        Data = imageAsBase64,
                        Width = picture.Width,
                        Height = picture.Height
                    });
                }
            }
            else if (this.Shape.TextBody != null)
            {
                // Create a textbody element.
                TextBodyElement textBodyElement = this.AppendElement<TextBodyElement>(TextBodyElement.ELEMENT_NAME);
                textBodyElement.Parent = this;

                // Load the textbody element from the Syncfusion textbody.
                textBodyElement.Load(this.Shape.TextBody);
            }

            // Apply any additional attributes provided by the end user.
            this.ApplyAttributes(this.ElementAttributes);

            this.Update();

            return this;
        }
    }
}
