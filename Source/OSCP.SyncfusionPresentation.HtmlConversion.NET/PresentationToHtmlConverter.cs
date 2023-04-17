using Syncfusion.Drawing;
using Syncfusion.Licensing;
using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    public class PresentationToHtmlConverter
    {
        private IPresentation PresentationDocument { get; set; }
        private HtmlDocument HtmlDocument { get; set; }

        public PresentationToHtmlConverter()
        {
        }

        public string Convert(IPresentation presentation)
        {
            this.PresentationDocument = presentation;

            this.HtmlDocument = new HtmlDocument();

            // Loop over all the slides.
            foreach (ISlide slide in this.PresentationDocument.Slides)
            {
                SlideElement slideElement = this.HtmlDocument.AddSlide();
                slideElement.Load(slide);
            }

            return this.HtmlDocument.OuterHtml;
        }
    }
}
