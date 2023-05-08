using OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml;
using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    /// <summary>
    /// Convert HTML to a Syncfusion.Presentaton.IPresentation object.
    /// </summary>
    internal class PptxDocument
    {
        /// <summary>
        /// Settings for the conversion of HTML to a Syncfusion.Presentaton.IPresentation object.
        /// </summary>
        internal static HtmlToPresentationConverterSettings Settings { get; private set; }

        /// <summary>
        /// Syncfusion.Presentaton.IPresentation object that is created.
        /// </summary>
        public IPresentation IPresentation { get; private set; }

        /// <summary>
        /// XmlDocument that represents the HTML presentation document.
        /// </summary>
        private XmlDocument XmlDocument { get; set; }

        /// <summary>
        /// List of Base64 Images.
        /// </summary>
        internal static List<Base64Image> Base64Images { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        internal PptxDocument(HtmlToPresentationConverterSettings settings)
        {
            PptxDocument.Settings = settings;

            this.XmlDocument = new XmlDocument();
        }

        /// <summary>
        /// Load the HTML presentation document.
        /// </summary>
        /// <param name="html">HTML that contains the presentation document.</param>
        internal void Load(string html)
        {
            // Load the HTML into the XmlDocument.
            this.XmlDocument.LoadXml(html);

            // Create an instance of a Syncfusion.Presentaton.IPresentation.
            this.IPresentation = Presentation.Create();

            // Locate the XmlNodes that are an HTMLDivElement with the pptx-slide class.
            // Each HTMLDivElement is an HTML presentation slide.
            XmlNodeList slides = this.XmlDocument.SelectNodes($"//div[contains(@class, '{PptxDocument.Settings.CssClass.Slides}')]/div[contains(@class, '{PptxDocument.Settings.CssClass.Slide}')]");

            // Loop over all the slides in the document.
            foreach (XmlNode slideNode in slides) 
            {
                // Create and load a slide.
                SlidePart slidePart = new SlidePart(this);
                slidePart.Load(slideNode);
            }
        }

        /// <summary>
        /// Load the HTML presentation document with an external list of Base 64 images.
        /// </summary>
        /// <param name="html">HTML that contains the presentation document.</param>
        /// <param name="base64Images">Base64 images for the HTML document.</param>
        internal void Load(string html, List<Base64Image> base64Images)
        {
            // Store the Base64 images.
            PptxDocument.Base64Images = base64Images;

            // Load the HTML.
            this.Load(html);
        }
    }
}
