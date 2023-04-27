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
    /// Wrapper for an HTMLDivElement that represents a Syncfusion.Presentaton.ITextBody object.
    /// </summary>
    internal class TextBodyPart : PartObject
    {
        /// <summary>
        /// Parent PartObject with a Syncfusion.Presentaton.IITextBody.
        /// </summary>
        internal IPartWithTextBody PartWithTextBody { set; get; }

        /// <summary>
        /// Get the ITextBody object.
        /// </summary>
        internal ITextBody ITextBody 
        { 
            get
            {
                return this.PartWithTextBody.ITextBody;
            }
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="partWithTextBody">Parent PartObject with a Syncfusion.Presentaton.IITextBody.</param>
        internal TextBodyPart(IPartWithTextBody partWithTextBody)
        {
            this.PartWithTextBody = partWithTextBody;
        }

        /// <summary>
        /// Load the TextBodyPart.
        /// </summary>
        /// <param name="textBodyNode">XmlNode that represents the HTMLDivElement.</param>
        internal virtual void Load(XmlNode textBodyNode)
        {
            this.Node = textBodyNode;

            string verticalAlign = this.Css("vertical-align");

            this.PartWithTextBody.ITextBody.VerticalAlignment = verticalAlign == "bottom"
                ? VerticalAlignmentType.Bottom
                : (
                    verticalAlign == "middle"
                        ? VerticalAlignmentType.Middle
                        : VerticalAlignmentType.Top
                );

            // Loop over all the child elements in the textbody node.
            foreach (XmlNode childNode in textBodyNode.ChildNodes)
            {
                // Child element is a paragraph.
                if (childNode.Name == "p")
                {
                    // Add a paragraph.
                    ParagraphPart paragraphPart = new ParagraphPart(this);
                    // Load the paragraph.
                    paragraphPart.Load(childNode);
                }
                // Child element is an unordered list.
                else if (childNode.Name == "ul")
                {
                    // Add an unordered list.
                    UnorderedListPart unorderedListPart = new UnorderedListPart(this);
                    // Load the unordered list.
                    unorderedListPart.Load(childNode, 1);
                }
                // Child element is an ordered list.
                else if (childNode.Name == "ol")
                {
                    // Add an ordered list.
                    OrderedListPart orderedListPart = new OrderedListPart(this);
                    // Load the ordered list.
                    orderedListPart.Load(childNode, 1);
                }
            }
        }
    }
}
