using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal class OrderedListElement : ListElement
    {
        public OrderedListElement(XmlNode node) : base(node)
        {
            this.AddClass("pptx-ordered-list");
        }

        internal override ListElement Load(IParagraph paragraph)
        {
            string listStyleType = string.Empty;
            
            switch (paragraph.ListFormat.NumberStyle)
            {
                case NumberedListStyle.AlphaUcPeriod:
                    listStyleType = "upper-alpha";
                    break;
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
