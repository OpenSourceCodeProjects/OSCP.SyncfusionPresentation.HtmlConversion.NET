using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class ParagraphElement : HtmlElement
    {
        internal const string ELEMENT_NAME = "p";

        internal IParagraph IParagraph { get; set; }
        internal TextBodyElement Parent { get; set; }

        public ParagraphElement(XmlNode node) : base(node)
        {
            this.AddClass(HtmlDocument.Settings.CssClass.Paragraph);
        }

        internal ParagraphElement Load(IParagraph paragraph)
        {
            this.IParagraph = paragraph;

            this.Css("display", "block")
                .Css("font-family", $"'{paragraph.Font.FontName}'")
                .Css("font-size", $"{paragraph.Font.FontSize}px")
                .Css("text-align", paragraph.HorizontalAlignment.ToString().ToLower());

            if (paragraph.TextParts.Count > 0)
            {
                foreach (ITextPart textPart in paragraph.TextParts)
                {
                    TextPartElement textPartElement = this.AppendElement<TextPartElement>(TextPartElement.ELEMENT_NAME);
                    textPartElement.Load(textPart);
                }
            }
            else
            {
                TextPartElement textPartElement = this.AppendElement<TextPartElement>(TextPartElement.ELEMENT_NAME);
                textPartElement.Empty();
            }

            // Apply any additional attributes provided by the end user.
            this.ApplyAttributes(HtmlDocument.Settings.ElementAttributes.Paragraph);

            this.Update();

            return this;
        }
    }
}
