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
            this.AddClass("pptx-unordered-list");
        }

        internal override ListElement Load(IParagraph paragraph)
        {
            //string listStyleType = string.Empty;

            //switch (paragraph.ListFormat.BulletCharacter)
            //{
            //    case '•':
            //        listStyleType = "disc";
            //        break;
            //    case 'o':
            //        listStyleType = "circle";
            //        break;
            //    default:
            //        if (paragraph.ListFormat.FontName == "Wingdings")
            //        {
            //            WINDINGS_DEC_TO_UNICODE_HEX.TryGetValue(Convert.ToInt32(paragraph.ListFormat.BulletCharacter), out listStyleType);
            //        }
            //        if (listStyleType == null || listStyleType.Length == 0)
            //        {
            //            listStyleType = $"\"{paragraph.ListFormat.BulletCharacter}\"";
            //        }

            //        this.Attr("data-font", paragraph.ListFormat.FontName);

            //        break;
            //}

            //this.Css("list-style-type", listStyleType);

            this.Attr("data-font", paragraph.ListFormat.FontName);
            this.Css("list-style-type",
                BulletCharacterConversion.BulletCharacterToListStyleType(
                    paragraph.ListFormat.FontName,
                    paragraph.ListFormat.BulletCharacter)
            );

            return base.Load(paragraph);
        }
    }
}
