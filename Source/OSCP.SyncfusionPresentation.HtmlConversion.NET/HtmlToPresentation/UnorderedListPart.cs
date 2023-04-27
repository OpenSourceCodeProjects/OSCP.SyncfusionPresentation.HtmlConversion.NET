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
    /// Convert an HTMLUListElement to a Syncfusion.Presentaton.IParagraph object
    /// with a ListFormat.Type == ListType.Bulleted.
    /// </summary>
    internal class UnorderedListPart : ListPart
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="textBodyPart">Parent TextBodyPart object.</param>
        internal UnorderedListPart(TextBodyPart textBodyPart) : base(textBodyPart)
        {
        }

        /// <summary>
        /// Load the unordered list content.
        /// </summary>
        /// <param name="unorderedListNode">XmlNode that represents the HTMLUListElement.</param>
        internal void Load(XmlNode unorderedListNode, int indentLevel)
        {
            this.Node = unorderedListNode;
            this.IndentLevel = indentLevel;

            ParagraphPart paragraphPart = null;

            string fontName = this.Attr("data-font");
            string listStyleType = this.Css("list-style-type");

            // Loop over the child items in the list.
            foreach (XmlNode xmlNode in unorderedListNode.ChildNodes)
            {
                paragraphPart = new ParagraphPart(this.TextBodyPart);
                paragraphPart.AddChildNode += this.OnAddChildNode;
                paragraphPart.Load(xmlNode);
                paragraphPart.IParagraph.ListFormat.Type = ListType.Bulleted;
                paragraphPart.IParagraph.ListFormat.BulletCharacter = BulletCharacterConversion.ListStyleTypeToBulletCharacter(fontName, listStyleType);
                paragraphPart.IParagraph.ListFormat.FontName = fontName;
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
                OrderedListPart orderedListPart = new OrderedListPart(this.TextBodyPart);
                orderedListPart.Load(e.XmlNode, this.IndentLevel + 1);
            }
            // The child node is an unordered list.
            else if (e.XmlNode.Name == "ul")
            {
                // Add and load an unordered list.
                UnorderedListPart unorderedListPart = new UnorderedListPart(this.TextBodyPart);
                unorderedListPart.Load(e.XmlNode, this.IndentLevel + 1);
            }
        }
    }
}
