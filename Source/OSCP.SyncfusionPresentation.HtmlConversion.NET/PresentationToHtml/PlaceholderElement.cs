using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class PlaceholderElement : AutoShapeElement
    {
        public PlaceholderElement(XmlNode node) : base(node)
        {
            this.CssClassSettings = HtmlDocument.Settings.CssClass.Placeholder;
        }
    }
}
