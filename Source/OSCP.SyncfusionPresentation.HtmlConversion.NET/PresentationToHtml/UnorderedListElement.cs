using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class UnorderedListElement : ListElement
    {
        internal const string ELEMENT_NAME = "ul";

        public UnorderedListElement(XmlNode node) : base(node)
        {
            this.AddClass(HtmlDocument.Settings.CssClass.UnorderedList);
        }

        internal override ListElement Load(IParagraph paragraph)
        {
            this.Attr("data-font", paragraph.ListFormat.FontName);
            this.Css("list-style-type",
                BulletCharacterConversion.BulletCharacterToListStyleType(
                    paragraph.ListFormat.FontName,
                    paragraph.ListFormat.BulletCharacter)
            );

            // Apply any additional attributes provided by the end user.
            this.ApplyAttributes(HtmlDocument.Settings.ElementAttributes.UnorderedList);

            return base.Load(paragraph);
        }
    }
}
