using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class HtmlDocument
    {
        internal static ConverterSettings Settings { get; set; }

        private XmlDocument XmlDocument { get; set; }
        private HtmlElement DocumentElement { get; set; }
        private HtmlElement SlidesElement { get; set; }

        internal HtmlDocument(ConverterSettings settings)
        {
            HtmlDocument.Settings = settings;

            this.XmlDocument = new XmlDocument();

            this.DocumentElement = this.AddElement("div");
            this.DocumentElement.AddClass(settings.CssClass.Document).Update();

            this.SlidesElement = this.DocumentElement.AppendElement<HtmlElement>("div");
            this.SlidesElement.AddClass(settings.CssClass.Slides).Update();
        }

        internal void Load(IPresentation presentation)
        {
            // Loop over all the slides.
            foreach (ISlide slide in presentation.Slides)
            {
                SlideElement slideElement = this.AddSlide();
                slideElement.Load(slide);
            }
        }

        internal SlideElement AddSlide()
        {
            SlideElement slideElement = this.SlidesElement.AppendElement<SlideElement>(SlideElement.ELEMENT_NAME);
            slideElement.Parent = this.SlidesElement;
            return slideElement;
        }

        private HtmlElement AddElement(string name)
        {
            XmlNode node = this.XmlDocument.CreateElement(name);
            (node as XmlElement).IsEmpty = false;
            this.XmlDocument.AppendChild(node);
            HtmlElement htmlElement = new HtmlElement(node);

            return htmlElement;
        }

        internal string OuterHtml
        {
            get
            {
                return this.XmlDocument.OuterXml;
            }
        }
    }
}
