using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal abstract class SlideItemElement : HtmlElement
    {
        internal SlideElement Parent { get; set; }

        public SlideItemElement(XmlNode node) : base(node)
        {
            this.AddClass(HtmlDocument.Settings.CssClass.SlideItem);
        }

        internal SlideItemElement PositionElement(ISlideItem slideItem)
        {
            this.Css("position", "absolute")
                .Css("display", "table")
                .Css("top", $"{slideItem.Top}px")
                .Css("left", $"{slideItem.Left}px")
                .Css("height", $"{slideItem.Height}px")
                .Css("width", $"{slideItem.Width}px");

            return this;
        }
    }
}