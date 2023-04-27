using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal abstract class ListElement : HtmlElement
    {
        internal IParagraph Paragraph { get; set; }
        internal TextBodyElement Parent { get; set; }
        protected List<HtmlElement> ListItems { get; set; }

        public ListElement(XmlNode node) : base(node)
        {
            this.ListItems = new List<HtmlElement>();
        }

        internal virtual ListElement Load(IParagraph paragraph)
        {
            this.Paragraph = paragraph;

            this.Css("display", "block")
                .Css("font-family", $"'{paragraph.Font.FontName}'")
                .Css("font-size", $"{paragraph.Font.FontSize}px")
                .Css("text-align", paragraph.HorizontalAlignment.ToString().ToLower());

            this.Update();

            return this;
        }

        internal ListElement AddListItem(IParagraph paragraph)
        {
            HtmlElement listItemElement = this.AppendElement<HtmlElement>("li");
            this.ListItems.Add(listItemElement);

            if (paragraph.TextParts.Count > 0)
            {
                foreach (ITextPart textPart in paragraph.TextParts)
                {
                    TextPartElement textPartElement = listItemElement.AppendElement<TextPartElement>(TextPartElement.ELEMENT_NAME);
                    textPartElement.Load(textPart);
                }
            }
            else
            {
                TextPartElement textPartElement = listItemElement.AppendElement<TextPartElement>(TextPartElement.ELEMENT_NAME);
                textPartElement.Empty();
            }

            return this;
        }

        internal T AddListToListItem<T>(int listItemIndex = -1)
        {
            HtmlElement listItemElement = this.ListItems[listItemIndex > -1 ? listItemIndex : (this.ListItems.Count - 1)];

            string tag = typeof(T) == typeof(UnorderedListElement) 
                ? UnorderedListElement.ELEMENT_NAME 
                : OrderedListElement.ELEMENT_NAME;

            T listElement = listItemElement.AppendElement<T>(tag);

            return listElement;
        }
    }
}
