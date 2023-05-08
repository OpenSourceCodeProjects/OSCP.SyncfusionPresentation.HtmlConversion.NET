using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    internal class AutoShapePart : SlideItemPart
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="slidePart">Parent SlidePart.</param>
        internal AutoShapePart(SlidePart slidePart) : base(slidePart)
        {
        }

        /// <summary>
        /// Load the auto shape.
        /// </summary>
        /// <param name="shapeNode">XmlNode that represents the HTMLDivElement.</param>
        internal override void Load(XmlNode shapeNode)
        {
            base.Load(shapeNode);

            // There is a child element and it has a TextBody CSS class.
            if (shapeNode.HasClass(PptxDocument.Settings.CssClass.Picture) == true)
            {
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
            else
            {
                // Get the child div element.
                XmlNode childNode = shapeNode.SelectSingleNode("div");

                if (childNode != null && childNode.HasClass(PptxDocument.Settings.CssClass.TextBody) == true)
                {
                    // Initialize the shape as a text box.
                    this.InitIShapeAsTextBox();

                    // Add a text body.
                    TextBodyPart textBodyPart = new TextBodyPart(this.IShape.TextBody);
                    // Load the text body.
                    textBodyPart.Load(childNode);
                }
            }
        }

        /// <summary>
        /// Initialize the shape as a text box.
        /// </summary>
        private void InitIShapeAsTextBox()
        {
            // The Syncfusion.Presentation.IShape has not been created yet.
            if (this.IShape == null)
            {
                // Create an instance of the Syncfusion.Presentation.IShape as a text box.
                this.IShape = this.Parent.ISlide.AddTextBox(this.Left, this.Top, this.Width, this.Height);
            }
        }
    }
}
