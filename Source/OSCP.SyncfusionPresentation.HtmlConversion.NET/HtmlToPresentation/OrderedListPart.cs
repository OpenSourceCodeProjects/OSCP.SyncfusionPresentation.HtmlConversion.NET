using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    /// <summary>
    /// Convert an HTMLOListElement to a Syncfusion.Presentaton.IParagraph object
    /// with a ListFormat.Type == ListType.Numbered.
    /// </summary>
    internal class OrderedListPart : ListPart
    {
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
            // Store parameter values passed in.
            this.Node = orderedListNode;
            this.IndentLevel = indentLevel;

            ParagraphPart paragraphPart = null;

            // Loop over the child items in the list.
            foreach (XmlNode xmlNode in orderedListNode.ChildNodes)
            {
                // Add a ParagraphPart.
                paragraphPart = new ParagraphPart(this.ShapePart);
                // Subscribe to the AddChildNode event.
                paragraphPart.AddChildNode += this.OnAddChildNode;
                // Load the paragraph.
                paragraphPart.Load(xmlNode);
                // Set the paragraph to be a numbered list.
                paragraphPart.IParagraph.ListFormat.Type = ListType.Numbered;
                // Set the number style.
                paragraphPart.IParagraph.ListFormat.NumberStyle = this.GetNumberedListStyle();
                // Set the indent level.
                paragraphPart.IParagraph.IndentLevelNumber = indentLevel;
            }
        }

        /// <summary>
        /// The paragraph is adding a child node.
        /// </summary>
        /// <param name="e"></param>
        internal void OnAddChildNode(AddChildNodeArgs e)
        {
            // The child node is an ordered list.
            if (e.XmlNode.Name == "ol")
            {
                // Add and load an ordered list.
                OrderedListPart orderedListPart = new OrderedListPart(this.ShapePart);
                orderedListPart.Load(e.XmlNode, this.IndentLevel + 1);
            }
            // The child node is an unordered list.
            else if (e.XmlNode.Name == "ul")
            {
                // Add and load an unordered list.
                UnorderedListPart unorderedListPart = new UnorderedListPart(this.ShapePart);
                unorderedListPart.Load(e.XmlNode, this.IndentLevel + 1);
            }
        }

        /// <summary>
        /// Get the number style.
        /// </summary>
        private NumberedListStyle GetNumberedListStyle()
        {
            // Default to number.
            NumberedListStyle numberedListStyle = NumberedListStyle.ArabicPeriod;

            // Get the list style type.
            string listStyleType = this.Css("list-style-type");

            switch (listStyleType)
            {
                // Upper case alpha characters.
                case "upper-alpha":
                    numberedListStyle = NumberedListStyle.AlphaUcPeriod;
                    break;
                // Lower case alpha characters.
                case "lower-alpha":
                    numberedListStyle = NumberedListStyle.AlphaLcPeriod;
                    break;
                // Upper case roman numerals.
                case "upper-roman":
                    numberedListStyle = NumberedListStyle.RomanUcPeriod;
                    break;
                // Lower case roman numerals.
                case "lower-roman":
                    numberedListStyle = NumberedListStyle.RomanLcPeriod;
                    break;
            }

            return numberedListStyle;
        }
    }
}
