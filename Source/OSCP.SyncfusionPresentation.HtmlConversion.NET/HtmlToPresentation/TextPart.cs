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
    /// Convert an HTMLSpanElement to a Syncfusion.Presentaton.ITextPart object.
    /// </summary>
    internal class TextPart : PartObject
    {
        /// <summary>
        /// Parent ParagraphPart object.
        /// </summary>
        internal ParagraphPart ParagraphPart { get; set; }

        /// <summary>
        /// Syncfusion.Presentaton.ITextPart object that is created.
        /// </summary>
        internal ITextPart ITextPart { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="paragraphPart">Parent ParagraphPart object.</param>
        internal TextPart(ParagraphPart paragraphPart)
        {
            this.ParagraphPart = paragraphPart;
        }

        /// <summary>
        /// Load the textpart.
        /// </summary>
        /// <param name="paragraphNode">XmlNode that represents the HTMLSpanElement.</param>
        internal void Load(XmlNode spanNode)
        {
            this.Node = spanNode;

            // Create the Syncfusion.Presentaton.ITextPart object.
            this.ITextPart = this.ParagraphPart.IParagraph.AddTextPart();

            string styleValue = string.Empty;

            // Apply font family and size.
            this.ITextPart.Font.FontName = this.Css("font-family");
            double fontSize = 0;
            if (this.TryGetCssValue("font-size", out fontSize) == true) this.ITextPart.Font.FontSize = (float)fontSize;

            // Apply text color.
            if ((styleValue = this.Css("color")).Length > 0)
            {
                ColorObject colorObject = ColorExtensions.FromCss(styleValue);
                if (colorObject != null)
                {
                    this.ITextPart.Font.Color = colorObject;
                }
            }

            // Apply bold.
            styleValue = this.Css("font-weight");
            if (styleValue == "bold") this.ITextPart.Font.Bold = true;

            // Apply italic.
            styleValue = this.Css("font-style");
            if (styleValue == "italic") this.ITextPart.Font.Italic = true;

            // Apply underline.
            styleValue = this.Css("text-decoration-line");
            if (styleValue == "underline")
            {
                // Apply the underline style.
                string underlineValue = this.Css("text-decoration-style");
                if (underlineValue == "dashed") this.ITextPart.Font.Underline = TextUnderlineType.Dash;
                else if (underlineValue == "dotted") this.ITextPart.Font.Underline = TextUnderlineType.Dotted;
                else if (underlineValue == "wavy") this.ITextPart.Font.Underline = TextUnderlineType.Wave;
                else this.ITextPart.Font.Underline = TextUnderlineType.Single;

                // Apply the underline color.
                if ((underlineValue = this.Css("text-decoration-color")).Length > 0)
                {
                    this.ITextPart.UnderlineColor = ColorExtensions.FromCss(underlineValue);
                }
                else
                {
                    this.ITextPart.UnderlineColor = ColorObject.Black;
                }
            }
            // Apply strikethrough.
            else if (styleValue == "line-through")
            {
                string lineThroughStyle = this.Css("text-decoration-style");
                this.ITextPart.Font.StrikeType = lineThroughStyle == "double"
                    ? TextStrikethroughType.Double
                    : TextStrikethroughType.Single;
            }

            // Apply All Caps.
            if ((styleValue = this.Css("text-transform")) == "uppercase")
            {
                this.ITextPart.Font.CapsType = TextCapsType.All;
            }

            // Apply Small Caps.
            if ((styleValue = this.Css("font-variant-caps")) == "small-caps")
            {
                this.ITextPart.Font.CapsType = TextCapsType.Small;
            }

            // Apply Highlight.
            if ((styleValue = this.Css("background-color")).Length > 0)
            {
                this.ITextPart.Font.HighlightColor = ColorExtensions.FromCss(styleValue);
            }

            // The HTMLSpanElement has a child element.
            if (this.Node.HasChildNodes)
            {
                // The child element is a sub (subscript) HTMLElement.
                if (this.Node.FirstChild.Name == "sub")
                {
                    this.ITextPart.Font.Subscript = true;
                }
                // The child element is a sup (superscript) HTMLElement.
                else if (this.Node.FirstChild.Name == "sup")
                {
                    this.ITextPart.Font.Superscript = true;
                }
            }

            // Apply text.
            this.ITextPart.Text = this.Node.InnerText;
        }
    }
}
