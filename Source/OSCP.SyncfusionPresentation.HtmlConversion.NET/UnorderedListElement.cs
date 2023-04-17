using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal class UnorderedListElement : ListElement
    {
        // http://www.alanwood.net/demos/wingdings.html
        private static readonly Dictionary<int, string> WINDINGS_DEC_TO_UNICODE_HEX = new Dictionary<int, string>()
        {
            { 113, "\"\\2751\"" },  // boxshadowdwn
            { 118, "\"\\2756\"" },  // xrhombus
            { 167, "square" },      // square4
            { 216, "\"\\2B9A\"" },  // head2right
            { 252, "\"\\2714\"" }   // checkbld
        };

        public UnorderedListElement(XmlNode node) : base(node)
        {
            this.AddClass("pptx-unordered-list");
        }

        internal override ListElement Load(IParagraph paragraph)
        {
            string listStyleType = string.Empty;

            switch (paragraph.ListFormat.BulletCharacter)
            {
                case '•':
                    listStyleType = "disc";
                    break;
                case 'o':
                    listStyleType = "circle";
                    break;
                default:
                    if (paragraph.ListFormat.FontName == "Wingdings")
                    {
                        WINDINGS_DEC_TO_UNICODE_HEX.TryGetValue(Convert.ToInt32(paragraph.ListFormat.BulletCharacter), out listStyleType);
                    }
                    if (listStyleType == null || listStyleType.Length == 0)
                    {
                        listStyleType = $"\"{paragraph.ListFormat.BulletCharacter}\"";
                    }

                    this.Attr("data-font", paragraph.ListFormat.FontName);

                    break;
            }

            this.Css("list-style-type", listStyleType);

            return base.Load(paragraph);
        }
    }
}
