using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal class TextPartElement : HtmlElement
    {
        internal const string ELEMENT_NAME = "span";

        internal ITextPart TextPart { get; set; }
        internal HtmlElement Parent { get; set; }

        public TextPartElement(XmlNode node) : base(node)
        {
            this.AddClass("pptx-textpart");
        }

        internal TextPartElement Empty()
        {
            this.Text("&nbsp;");

            return this;
        }

        internal TextPartElement Load(ITextPart textPart)
        {
            this.TextPart = textPart;

            HtmlElement htmlElement = this;

            // This is a superscript.
            if (textPart.Font.Superscript == true)
            {
                htmlElement = this.AddElement<HtmlElement>("sup");
            }
            else if (textPart.Font.Subscript == true) 
            {
                htmlElement = this.AddElement<HtmlElement>("sub");
            }

            htmlElement.Css("color", $"rgb({textPart.Font.Color.R},{textPart.Font.Color.G},{textPart.Font.Color.B})");

            if (textPart.Font.Bold == true) htmlElement.Css("font-weight", "bold");
            if (textPart.Font.Italic == true) htmlElement.Css("font-style", "italic");
            if (textPart.Font.Underline != TextUnderlineType.None)
            {
                htmlElement.Css("text-decoration-line", "underline")
                    .Css("text-decoration-color", $"rgb({textPart.UnderlineColor.R},{textPart.UnderlineColor.G},{textPart.UnderlineColor.B})")
                    .Css("text-decoration-style",
                        textPart.Font.Underline == TextUnderlineType.Wave
                            ? "wavy"
                            : (
                                textPart.Font.Underline == TextUnderlineType.Dash
                                    ? "dashed"
                                    : (
                                        textPart.Font.Underline == TextUnderlineType.Dotted
                                            ? "dotted"
                                            : "solid"
                                    )
                            )
                    );
            }
            if (textPart.Font.StrikeType != TextStrikethroughType.None) 
            {
                htmlElement.Css("text-decoration-line", "line-through");
                if (textPart.Font.StrikeType == TextStrikethroughType.Double) htmlElement.Css("text-decoration-style", "double");
            }

            if (textPart.Font.CapsType == TextCapsType.All) this.Css("text-transform", "uppercase");
            if (textPart.Font.CapsType == TextCapsType.Small) this.Css("font-variant-caps", "small-caps");

            htmlElement.Text(textPart.Text);
            htmlElement.Update();

            return this;
        }
    }
}
