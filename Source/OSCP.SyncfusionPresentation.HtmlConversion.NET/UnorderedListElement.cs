using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal class UnorderedListElement : HtmlElement
    {
        internal IParagraph Paragraph { get; set; }
        internal ShapeElement Parent { get; set; }

        public UnorderedListElement(XmlNode node) : base(node)
        {
            this.AddClass("pptx-unordered-list");
        }

        internal UnorderedListElement Load(IParagraph paragraph)
        {
            this.Paragraph = paragraph;

            string listStyleType = paragraph.ListFormat.BulletCharacter == '•'
                ? "disc"
                : $"{paragraph.ListFormat.BulletCharacter}";

            this.Css("display", "table-cell")
                .Css("font-family", $"'{paragraph.Font.FontName}'")
                .Css("font-size", $"{paragraph.Font.FontSize}px")
                .Css("vertical-align", this.Parent.Shape.TextBody.VerticalAlignment.ToString().ToLower())
                .Css("text-align", paragraph.HorizontalAlignment.ToString().ToLower())
                .Css("list-style-type", listStyleType);

            this.Update();

            return this;
        }

        internal UnorderedListElement AddListItem(string text)
        {
            HtmlElement listItemElement = this.AddElement<HtmlElement>("li");
            HtmlElement spanElement = listItemElement.AddElement<HtmlElement>("span");
            spanElement.Text(text);

            return this;
        }
    }
}
