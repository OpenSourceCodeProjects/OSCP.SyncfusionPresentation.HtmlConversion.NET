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
    internal class PptxDocument
    {
        public IPresentation PresentationDocument { get; private set; }

        private XmlDocument XmlDocument { get; set; }

        internal PptxDocument()
        {
            this.XmlDocument = new XmlDocument();
        }

        internal void Load(string html)
        {
            this.XmlDocument.LoadXml(html);

            this.PresentationDocument = Presentation.Create();

            // Get the slides.
            XmlNodeList slides = this.XmlDocument.SelectNodes("//div[contains(@class, 'pptx-slides')]/div[contains(@class, 'pptx-slide')]");

            foreach (XmlNode slideNode in slides) 
            {
                SlidePart slidePart = this.AddSlide();
                slidePart.Load(slideNode);
            }
        }

        internal SlidePart AddSlide()
        {
            SlidePart slidePart = new SlidePart();
            slidePart.Parent = this.PresentationDocument;
            return slidePart;
        }
    }
}
