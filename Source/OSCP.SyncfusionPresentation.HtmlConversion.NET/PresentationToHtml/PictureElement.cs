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
        }

        internal PictureElement Load(IPicture picture)
        {
            this.Picture = picture;

            // Position the element on the slide.
            this.PositionElement(picture);

            string imageType = picture.ImageFormat.ToString().ToLower();

            // Add the background image as a base64 string.
            string imageAsBase64 = System.Convert.ToBase64String(picture.ImageData);

            this.Css("background-image", $"url(data:image/{imageType};base64,{imageAsBase64})")
                .Css("background-size", $"{picture.Width}px {picture.Height}px");

            // Update the underlying XmlNode with the styles and classes.
            this.Update();

            return this;
        }
    }
}
