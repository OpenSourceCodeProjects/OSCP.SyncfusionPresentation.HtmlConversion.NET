using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    internal class ParagraphPart : ShapePart
    {
        internal IParagraph IParagraph { get; set; }

        internal ParagraphPart(SlidePart slidePart) : base(slidePart)
        {
        }

        internal override void Load(XmlNode shapeNode)
        {
            base.Load(shapeNode);

            this.IShape = this.SlidePart.ISlide.AddTextBox(this.Left, this.Top, this.Width, this.Height);
            this.IParagraph = this.IShape.TextBody.AddParagraph();

            this.IParagraph.Font.FontName = this.Css("font-family");
            double fontSize = 0;
            if (this.TryGetCssValue("font-size", out fontSize) == true) this.IParagraph.Font.FontSize = (float)fontSize;

            foreach (XmlNode spanNode in shapeNode.FirstChild.ChildNodes)
            {
                TextPart textPart = new TextPart(this);
                textPart.Load(spanNode);
            }
        }
    }
}
