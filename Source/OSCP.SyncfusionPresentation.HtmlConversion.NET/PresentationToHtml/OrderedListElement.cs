using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class OrderedListElement : ListElement
    {
        internal const string ELEMENT_NAME = "ol";

        public OrderedListElement(XmlNode node) : base(node)
        {
            this.AddClass(HtmlDocument.Settings.CssClass.OrderedList);
        }

        internal override ListElement Load(IParagraph paragraph)
        {
            string listStyleType = string.Empty;
            
            switch (paragraph.ListFormat.NumberStyle)
            {
                case NumberedListStyle.AlphaUcParenRight:
                case NumberedListStyle.AlphaUcPeriod:
                    listStyleType = "upper-alpha";
                    break;
                case NumberedListStyle.AlphaLcParenRight:
                case NumberedListStyle.AlphaLcPeriod:
                    listStyleType = "lower-alpha";
                    break;
                case NumberedListStyle.RomanUcPeriod:
                    listStyleType = "upper-roman";
                    break;
                case NumberedListStyle.RomanLcPeriod:
                    listStyleType = "lower-roman";
                    break;
                default:
                    listStyleType = "decimal";
                    break;
            }

            this.Css("list-style-type", listStyleType);

            return base.Load(paragraph);
        }
    }
}
