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
    internal class SlidePart : PartObject
    {
        internal IPresentation Parent { set; get; }
        internal ISlide Slide { get; set; }

        internal void Load(XmlNode slideNode)
        {
            this.Node = slideNode;

            // Adds a slide to the PowerPoint presentation
            this.Slide = this.Parent.Slides.Add();

            this.ApplyBackground(slideNode);
        }

        /// <summary>
        /// Determine whether the slide has a background and if so, then apply it.
        /// </summary>
        /// <param name="slideNode">HTML slide element.</param>
        private void ApplyBackground(XmlNode slideNode)
        {
            string styleValue = string.Empty;

            // Retrieve the background instance.
            IBackground background = this.Slide.Background;

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
                Match imageMatch = Regex.Match(styleValue, @"url\(data:image/[^;]*;base64,([^\)]*)\)");
                if (imageMatch.Success == true) 
                {
                    string base64ImageString = imageMatch.Groups[1].Value;
                    byte[] imageBytes = Convert.FromBase64String(base64ImageString).ToArray();
                    background.Fill.FillType = FillType.Picture;
                    background.Fill.PictureFill.ImageBytes = imageBytes;
                }
            }
        }
    }
}
