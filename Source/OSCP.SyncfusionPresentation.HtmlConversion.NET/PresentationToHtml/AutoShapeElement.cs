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

        public AutoShapeElement(XmlNode node) : base(node)
        {
            this.CssClassSettings = HtmlDocument.Settings.CssClass.AutoShape;
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
                // Cast the shape as a picture to get the image information.
                IPicture picture = shape as IPicture;

                // Get the image type.
                string imageType = picture.ImageFormat.ToString().ToLower();

                // Get the image as a base64 string.
                string imageAsBase64 = System.Convert.ToBase64String(picture.ImageData);

                // Set the background of this shape with the image.
                this.Css("display", "initial")
                    .Css("background-image", $"url(data:image/{imageType};base64,{imageAsBase64})")
                    .Css("background-size", $"{picture.Width}px {picture.Height}px")
                    .Css("background-repeat", "no-repeat");
            }
            else if (this.Shape.TextBody != null)
            {
                // Create a textbody element.
                TextBodyElement textBodyElement = this.AppendElement<TextBodyElement>(TextBodyElement.ELEMENT_NAME);
                textBodyElement.Parent = this;

                // Load the textbody element from the Syncfusion textbody.
                textBodyElement.Load(this.Shape.TextBody);
            }

            this.Update();

            return this;
        }
    }
}
