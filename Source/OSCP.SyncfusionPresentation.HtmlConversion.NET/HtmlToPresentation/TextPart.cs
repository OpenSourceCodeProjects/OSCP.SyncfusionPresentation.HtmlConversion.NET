using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    internal class TextPart : PartObject
    {
        internal ParagraphPart ParagraphPart { get; set; }
        internal ITextPart ITextPart { get; set; }

        internal TextPart(ParagraphPart paragraphPart)
        {
            this.ParagraphPart = paragraphPart;
        }

        internal void Load(XmlNode spanNode)
        {
            this.Node = spanNode;

            this.ITextPart = this.ParagraphPart.IParagraph.AddTextPart();

            string styleValue = string.Empty;

            this.ITextPart.Font.FontName = this.Css("font-family");
            double fontSize = 0;
            if (this.TryGetCssValue("font-size", out fontSize) == true) this.ITextPart.Font.FontSize = (float)fontSize;

            if ((styleValue = this.Css("color")).Length > 0)
            {
                ColorObject colorObject = ColorExtensions.FromCss(styleValue);
                if (colorObject != null) 
                {
                    this.ITextPart.Font.Color = colorObject;
                }
            }

            this.ITextPart.Text = this.Node.InnerText;
        }
    }
}
