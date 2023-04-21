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

            // Loop over the span elements within the paragraph.
            foreach (XmlNode spanNode in paragraphNode.ChildNodes)
            {
                // Create and load a text part.
                TextPart textPart = new TextPart(this);
                textPart.Load(spanNode);
            }
        }
    }
}
