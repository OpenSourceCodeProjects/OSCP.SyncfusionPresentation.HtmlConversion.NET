using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class PictureElement : SlideItemElement
    {
        internal const string ELEMENT_NAME = "img";

        internal IPicture Picture { get; set; }

        public PictureElement(XmlNode node) : base(node)
        {
            // Add the css class for a picture.
            this.AddClass(HtmlDocument.Settings.CssClass.Picture);
        }

        internal PictureElement Load(IPicture picture)
        {
            this.Picture = picture;

            // Position the element on the slide.
            this.PositionElement(picture);

            string imageType = picture.ImageFormat.ToString().ToLower();

            // Convert the image to a base64 string.
            string imageAsBase64 = System.Convert.ToBase64String(picture.ImageData);

            // Images are to be embedded in the HTML.
            if (HtmlDocument.Settings.ImageData.Location == ImageDataLocation.Embedded)
            {
                // Set the image as a background image for the element.
                this.Css("background-image", $"url(data:image/{imageType};base64,{imageAsBase64})")
                    .Css("background-size", $"{picture.Width}px {picture.Height}px");
            }
            // Images are to be stored as external image strings and referenced by the element
            // using a data-image-id attribute. For example, in a CSS file using the following
            // CSS:
            // img[data-image-id="image id here..."] {
            //      background-image: ...;
            //      backgound-size: ...;
            // }.
            else if (HtmlDocument.Settings.ImageData.Location == ImageDataLocation.Base64)
            {
                this.Css("display", "initial");

                string imageId = Guid.NewGuid().ToString();
                this.Attr("data-image-id", imageId);

                HtmlDocument.Settings.ImageData.Base64Images.Add(new Base64Image
                {
                    Id = imageId,
                    ElementName = PictureElement.ELEMENT_NAME,
                    ElementCssClass = HtmlDocument.Settings.CssClass.Picture,
                    ImageType = imageType,
                    Data = imageAsBase64,
                    Width = picture.Width,
                    Height = picture.Height
                });
            }

            // Apply any additional attributes provided by the end user.
            this.ApplyAttributes(HtmlDocument.Settings.ElementAttributes.Picture);

            // Update the underlying XmlNode with the styles and classes.
            this.Update();

            return this;
        }
    }
}
