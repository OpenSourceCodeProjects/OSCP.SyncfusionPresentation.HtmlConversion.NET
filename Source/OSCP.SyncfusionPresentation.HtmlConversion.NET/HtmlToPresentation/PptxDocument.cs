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
    /// Convert a HTML to a Syncfusion.Presentaton.IPresentation object.
    /// </summary>
    internal class PptxDocument
    {
        /// <summary>
        /// Syncfusion.Presentaton.IPresentation object that is created.
        /// </summary>
        public IPresentation IPresentation { get; private set; }

        /// <summary>
        /// XmlDocument that represents the HTML presentation document.
        /// </summary>
        private XmlDocument XmlDocument { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        internal PptxDocument()
        {
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
            XmlNodeList slides = this.XmlDocument.SelectNodes("//div[contains(@class, 'pptx-slides')]/div[contains(@class, 'pptx-slide')]");

            // Loop over all the slides in the document.
            foreach (XmlNode slideNode in slides) 
            {
                // Create and load a slide.
                SlidePart slidePart = new SlidePart(this);
                slidePart.Load(slideNode);
            }
        }
    }
}
