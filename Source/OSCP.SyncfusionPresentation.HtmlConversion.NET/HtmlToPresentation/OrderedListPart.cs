using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    internal class OrderedListPart : ListPart
    {
        internal int IndentLevel { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="shapePart">Parent ShapePart object.</param>
        internal OrderedListPart(ShapePart shapePart) : base(shapePart)
        {
        }

        /// <summary>
        /// Load the ordered list content.
        /// </summary>
        /// <param name="orderedListNode">XmlNode that represents the HTMLOListElement.</param>
        internal void Load(XmlNode orderedListNode, int indentLevel)
        {
            this.Node = orderedListNode;
            this.IndentLevel = indentLevel;

            ParagraphPart paragraphPart = null;

            // Loop over the child items in the list.
            foreach (XmlNode xmlNode in orderedListNode.ChildNodes)
            {
                paragraphPart = new ParagraphPart(this.ShapePart);
                paragraphPart.AddChildNode += this.OnAddChildNode;
                paragraphPart.Load(xmlNode);
                paragraphPart.IParagraph.ListFormat.Type = ListType.Numbered;
                paragraphPart.IParagraph.ListFormat.NumberStyle = this.GetNumberedListStyle();
                paragraphPart.IParagraph.IndentLevelNumber = indentLevel;
            }
        }

        internal void OnAddChildNode(AddChildNodeArgs e)
        {
            if (e.XmlNode.Name == "ol")
            {
                OrderedListPart orderedListPart = new OrderedListPart(this.ShapePart);
                orderedListPart.Load(e.XmlNode, this.IndentLevel + 1);
            }
        }

        private NumberedListStyle GetNumberedListStyle()
        {
            NumberedListStyle numberedListStyle = NumberedListStyle.ArabicPeriod;

            string listStyleType = this.Css("list-style-type");

            switch (listStyleType)
            {
                case "upper-alpha":
                    numberedListStyle = NumberedListStyle.AlphaUcPeriod;
                    break;
                case "lower-alpha":
                    numberedListStyle = NumberedListStyle.AlphaLcPeriod;
                    break;
                case "upper-roman":
                    numberedListStyle = NumberedListStyle.RomanUcPeriod;
                    break;
                case "lower-roman":
                    numberedListStyle = NumberedListStyle.RomanLcPeriod;
                    break;
            }

            return numberedListStyle;
        }
    }
}
