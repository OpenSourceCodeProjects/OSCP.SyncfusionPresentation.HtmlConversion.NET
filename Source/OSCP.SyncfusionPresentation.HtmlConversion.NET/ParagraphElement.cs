using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    public class ParagraphElement : HtmlElement
    {
        internal IParagraph Paragraph { get; set; }
        internal ShapeElement Parent { get; set; }

        public ParagraphElement(XmlNode node) : base(node)
        {
            this.AddClass("pptx-paragraph");
        }

        internal ParagraphElement Load(IParagraph paragraph)
        {
            this.Paragraph = paragraph;

            this.Css("display", "table-cell")
                .Css("font-family", $"'{paragraph.Font.FontName}'")
                .Css("font-size", $"{paragraph.Font.FontSize}px")
                .Css("vertical-align", this.Parent.Shape.TextBody.VerticalAlignment.ToString().ToLower())
                .Css("text-align", paragraph.HorizontalAlignment.ToString().ToLower());

            if (paragraph.TextParts.Count > 0)
            {
                foreach (ITextPart textPart in paragraph.TextParts)
                {
                    HtmlElement spanElement = this.AddElement<HtmlElement>("span");
                    spanElement.Css("color", $"rgb({textPart.Font.Color.R},{textPart.Font.Color.G},{textPart.Font.Color.B})");
                        
                    if (textPart.Font.Bold == true) spanElement.Css("font-weight", "bold");
                    if (textPart.Font.Italic == true) spanElement.Css("font-style", "italic");
                    if (textPart.Font.Underline != TextUnderlineType.None)
                    {
                        spanElement
                            .Css("text-decoration-line", "underline")
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
                    spanElement.Text(textPart.Text);
                    spanElement.Update();
                }
            }
            else
            {
                this.Text("&nbsp;");
            }

            this.Update();

            return this;
        }
    }
}
