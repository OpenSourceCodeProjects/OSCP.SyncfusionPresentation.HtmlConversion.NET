using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class TextPartElement : HtmlElement
    {
        internal const string ELEMENT_NAME = "span";

        internal ITextPart TextPart { get; set; }
        internal HtmlElement Parent { get; set; }

        public TextPartElement(XmlNode node) : base(node)
        {
            this.AddClass(HtmlDocument.Settings.CssClass.TextPart);
        }

        internal TextPartElement Empty()
        {
            this.Html("&nbsp;");

            return this;
        }

        /// <summary>
        /// Load the text part.
        /// </summary>
        /// <param name="textPart"></param>
        /// <returns></returns>
        internal TextPartElement Load(ITextPart textPart)
        {
            this.TextPart = textPart;

            HtmlElement htmlElement = this;

            // Text is superscript.
            if (textPart.Font.Superscript == true)
            {
                htmlElement = this.AppendElement<HtmlElement>("sup");
            }
            // Text is subscript.
            else if (textPart.Font.Subscript == true) 
            {
                htmlElement = this.AppendElement<HtmlElement>("sub");
            }

            htmlElement
                // Text color.
                .Css("color", $"rgb({textPart.Font.Color.R},{textPart.Font.Color.G},{textPart.Font.Color.B})")
                // Font family and size.
                .Css("font-family", $"'{textPart.Font.FontName}'")
                .Css("font-size", $"{textPart.Font.FontSize}px");
            
            // Text is bold.
            if (textPart.Font.Bold == true) htmlElement.Css("font-weight", "bold");

            // Text is italic.
            if (textPart.Font.Italic == true) htmlElement.Css("font-style", "italic");

            // Text is underlined.
            if (textPart.Font.Underline != TextUnderlineType.None)
            {
                htmlElement
                    // Underline the text.
                    .Css("text-decoration-line", "underline")
                    // Set the underline color.
                    .Css("text-decoration-color", $"rgb({textPart.UnderlineColor.R},{textPart.UnderlineColor.G},{textPart.UnderlineColor.B})")
                    // Set the underline format.
                    .Css("text-decoration-style",
                        // Underline is wavy.
                        textPart.Font.Underline == TextUnderlineType.Wave
                            ? "wavy"
                            : (
                                // Underline is dashed.
                                textPart.Font.Underline == TextUnderlineType.Dash
                                    ? "dashed"
                                    : (
                                        // Underline is dotted.
                                        textPart.Font.Underline == TextUnderlineType.Dotted
                                            ? "dotted"
                                            // Otherwise underline is solid.
                                            : "solid"
                                    )
                            )
                    );
            }

            // Strikethrough.
            if (textPart.Font.StrikeType != TextStrikethroughType.None) 
            {
                // Line through the text.
                htmlElement.Css("text-decoration-line", "line-through");

                // Double Strikethrough, apply double line through text.
                if (textPart.Font.StrikeType == TextStrikethroughType.Double) htmlElement.Css("text-decoration-style", "double");
            }

            // Text is uppercase.
            if (textPart.Font.CapsType == TextCapsType.All) this.Css("text-transform", "uppercase");

            // Text is small caps.
            if (textPart.Font.CapsType == TextCapsType.Small) this.Css("font-variant-caps", "small-caps");

            // Text is highlighted.
            if (textPart.Font.HighlightColor.SystemColor.Name != "ff000000")
            {
                htmlElement.Css("background-color", textPart.Font.HighlightColor.SystemColor.ToCss());
            }

            htmlElement.Text(textPart.Text);

            // Apply any additional attributes provided by the end user.
            this.ApplyAttributes(HtmlDocument.Settings.ElementAttributes.TextPart);

            htmlElement.Update();

            return this;
        }
    }
}
