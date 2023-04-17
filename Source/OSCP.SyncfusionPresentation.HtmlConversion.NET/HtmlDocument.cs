using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal class HtmlDocument
    {
        private XmlDocument XmlDocument { get; set; }
        private HtmlElement DocumentElement { get; set; }
        private HtmlElement SlidesElement { get; set; }

        internal HtmlDocument()
        {
            this.XmlDocument = new XmlDocument();

            this.DocumentElement = this.AddElement("div");
            this.DocumentElement.AddClass("pptx-document").Update();

            this.SlidesElement = this.DocumentElement.AddElement<HtmlElement>("div");
            this.SlidesElement.AddClass("pptx-slides").Update();
        }

        internal SlideElement AddSlide()
        {
            SlideElement slideElement = this.SlidesElement.AddElement<SlideElement>("div");
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
