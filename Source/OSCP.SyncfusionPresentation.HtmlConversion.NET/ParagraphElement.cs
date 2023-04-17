using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    public class ParagraphElement : HtmlElement
    {
        internal IParagraph Paragraph { get; set; }
        internal ShapeElement Parent { get; set; }

        public ParagraphElement(XmlNode node) : base(node)
        {
            this.AddClass("pptx-paragraph");
        }

        internal ParagraphElement Load(IParagraph paragraph)
        {
            this.Paragraph = paragraph;

            this.Css("display", "table-cell")
                .Css("font-family", $"'{paragraph.Font.FontName}'")
                .Css("font-size", $"{paragraph.Font.FontSize}px")
                .Css("vertical-align", this.Parent.Shape.TextBody.VerticalAlignment.ToString().ToLower())
                .Css("text-align", paragraph.HorizontalAlignment.ToString().ToLower());

            if (paragraph.TextParts.Count > 0)
            {
                foreach (ITextPart textPart in paragraph.TextParts)
                {
                    TextPartElement textPartElement = this.AddElement<TextPartElement>(TextPartElement.ELEMENT_NAME);
                    textPartElement.Load(textPart);
                }
            }
            else
            {
                TextPartElement textPartElement = this.AddElement<TextPartElement>(TextPartElement.ELEMENT_NAME);
                textPartElement.Empty();
            }

            this.Update();

            return this;
        }
    }
}
