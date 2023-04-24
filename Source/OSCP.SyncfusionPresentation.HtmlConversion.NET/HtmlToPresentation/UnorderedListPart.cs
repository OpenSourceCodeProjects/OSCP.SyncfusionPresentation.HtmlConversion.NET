using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    internal class UnorderedListPart : ListPart
    {
        internal int IndentLevel { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="shapePart">Parent ShapePart object.</param>
        internal UnorderedListPart(ShapePart shapePart) : base(shapePart)
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
                paragraphPart = new ParagraphPart(this.ShapePart);
                paragraphPart.AddChildNode += this.OnAddChildNode;
                paragraphPart.Load(xmlNode);
                paragraphPart.IParagraph.ListFormat.Type = ListType.Bulleted;
                paragraphPart.IParagraph.ListFormat.BulletCharacter = BulletCharacterConversion.ListStyleTypeToBulletCharacter(fontName, listStyleType);
                paragraphPart.IParagraph.ListFormat.FontName = fontName;
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
            else if (e.XmlNode.Name == "ul")
            {
                UnorderedListPart unorderedListPart = new UnorderedListPart(this.ShapePart);
                unorderedListPart.Load(e.XmlNode, this.IndentLevel + 1);
            }
        }
    }
}
