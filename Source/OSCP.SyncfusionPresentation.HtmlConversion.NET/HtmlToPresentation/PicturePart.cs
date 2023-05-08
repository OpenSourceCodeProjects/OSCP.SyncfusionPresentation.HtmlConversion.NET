using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    internal class PicturePart : SlideItemPart
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="slidePart">Parent SlidePart.</param>
        internal PicturePart(SlidePart slidePart) : base(slidePart)
        {
        }

        /// <summary>
        /// Load the picture.
        /// </summary>
        /// <param name="node">XmlNode that represents the HTMLImageElement.</param>
        internal override void Load(XmlNode node)
        {
            base.Load(node);

            string styleValue = string.Empty;

            byte[] imageBytes = null;

            // Background image.
            if ((styleValue = this.Css("background-image")).Length > 0)
            {
                // Get the base 64 string with the image.
                Match imageMatch = Regex.Match(styleValue, @"url\(data:image/[^;]*;base64,([^\)]*)\)");

                // A base 64 string for the image was found.
                if (imageMatch.Success == true)
                {
                    // Get the base64 string and apply the background image.
                    string base64ImageString = imageMatch.Groups[1].Value;
                    imageBytes = Convert.FromBase64String(base64ImageString).ToArray();
                }
            }
            else
            {
                // Check for an image id.
                string imageId = this.Attr("data-image-id");

                // There is an image id and there is a list of base 64 images.
                if (imageId.Length > 0 && PptxDocument.Base64Images != null)
                {
                    // Search for the base 64 image.
                    Base64Image base64Image = PptxDocument.Base64Images.Where(item => item.Id == imageId).FirstOrDefault();

                    // The base 64 image was found.
                    if (base64Image != null)
                    {
                        // Get the base64 string and apply the background image.
                        imageBytes = Convert.FromBase64String(base64Image.Data).ToArray();
                    }
                }
            }

            if (imageBytes != null)
            {
                using (MemoryStream stream = new MemoryStream(imageBytes))
                {
                    // Add the picture.
                    this.Parent.ISlide.Pictures.AddPicture(stream, this.Left, this.Top, this.Width, this.Height);
                }
            }
        }
    }
}
