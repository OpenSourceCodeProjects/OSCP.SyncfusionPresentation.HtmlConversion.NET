using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class LayoutSlideElement : SlideElement
    {
        public LayoutSlideElement(XmlNode node) : base(node)
        {
            this.CssClassSettings = HtmlDocument.Settings.CssClass.LayoutSlide;

            this.ElementAttributes = HtmlDocument.Settings.ElementAttributes.LayoutSlide;
        }
    }
}
