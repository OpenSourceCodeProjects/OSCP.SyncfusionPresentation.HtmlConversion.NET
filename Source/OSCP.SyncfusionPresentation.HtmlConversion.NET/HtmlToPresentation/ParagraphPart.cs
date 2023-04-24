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
    /// Convert an HTMLParagraphElement to a Syncfusion.Presentaton.IParagraph object.
    /// </summary>
    internal class ParagraphPart : PartObject
    {
        internal event AddChildNodeDelegate AddChildNode;

        /// <summary>
        /// Parent ShapePart object.
        /// </summary>
        internal ShapePart ShapePart { get; set; }

        /// <summary>
        /// Syncfusion.Presentaton.IParagraph object that is created.
        /// </summary>
        internal IParagraph IParagraph { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="shapePart">Parent ShapePart object.</param>
        internal ParagraphPart(ShapePart shapePart)
        {
            this.ShapePart = shapePart;
        }

        /// <summary>
        /// Load the paragraph.
        /// </summary>
        /// <param name="paragraphNode">XmlNode that represents the HTMLParagraphElement.</param>
        internal void Load(XmlNode paragraphNode)
        {
            this.Node = paragraphNode;

            string styleValue = string.Empty;

            // Create the Syncfusion.Presentaton.IParagraph object.
            this.IParagraph = this.ShapePart.IShape.TextBody.AddParagraph();

            // Apply font family and size.
            this.IParagraph.Font.FontName = this.Css("font-family");
            double fontSize = 0;
            if (this.TryGetCssValue("font-size", out fontSize) == true) this.IParagraph.Font.FontSize = (float)fontSize;

            // Apply vertical alignment.
            if ((styleValue = this.Css("vertical-align")).Length > 0)
            {
                if (styleValue == "top") this.ShapePart.IShape.TextBody.VerticalAlignment = VerticalAlignmentType.Top;
                else if (styleValue == "middle") this.ShapePart.IShape.TextBody.VerticalAlignment = VerticalAlignmentType.Middle;
                else if (styleValue == "bottom") this.ShapePart.IShape.TextBody.VerticalAlignment = VerticalAlignmentType.Bottom;
            }

            // Apply horizontal alignment.
            if ((styleValue = this.Css("text-align")).Length > 0)
            {
                if (styleValue == "left") this.IParagraph.HorizontalAlignment = HorizontalAlignmentType.Left;
                else if (styleValue == "center") this.IParagraph.HorizontalAlignment = HorizontalAlignmentType.Center;
                else if (styleValue == "right") this.IParagraph.HorizontalAlignment = HorizontalAlignmentType.Right;
            }
        //}

        ///// <summary>
        ///// Load the content for the paragraph. The default behavior assumes the
        ///// Syncfusion.Presentation.IParagraph.ListFormat.Type == ListType.None.
        ///// Override to load content for other ListType formats.
        ///// </summary>
        //internal virtual void LoadContent()
        //{
            // Loop over the span elements within the paragraph.
            foreach (XmlNode childNode in this.Node.ChildNodes)
            {
                if (childNode.Name == "span")
                {
                    // Create and load a text part.
                    TextPart textPart = new TextPart(this);
                    textPart.Load(childNode);
                }

                if (this.AddChildNode != null)
                {
                    this.AddChildNode(new AddChildNodeArgs { XmlNode = childNode });
                }
            }
        }
    }
}
