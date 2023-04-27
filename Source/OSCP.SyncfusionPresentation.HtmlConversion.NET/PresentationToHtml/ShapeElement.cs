using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class ShapeElement : HtmlElement
    {
        internal const string ELEMENT_NAME = "div";

        internal IShape Shape { get; set; }
        internal SlideElement Parent { get; set; }

        public ShapeElement(XmlNode node) : base(node)
        {
            this.AddClass(HtmlDocument.Settings.CssClass.Shape);
        }

        internal ShapeElement Load(IShape shape)
        {
            this.Shape = shape;

            this.Css("position", "absolute")
                .Css("display", "table")
                .Css("top", $"{shape.Top}px")
                .Css("left", $"{shape.Left}px")
                .Css("height", $"{shape.Height}px")
                .Css("width", $"{shape.Width}px");

            if (this.Shape.TextBody != null)
            {
                //string alignItems = this.Shape.TextBody.VerticalAlignment == VerticalAlignmentType.Top
                //    ? "start"
                //    : (
                //        this.Shape.TextBody.VerticalAlignment == VerticalAlignmentType.Middle
                //            ? "center"
                //            : "end"
                //    );

                //this.Css("align-items", alignItems);

                // Create a textbody element.
                TextBodyElement textBodyElement = this.AppendElement<TextBodyElement>(TextBodyElement.ELEMENT_NAME);
                textBodyElement.Parent = this;

                // Load the textbody element from the Syncfusion textbody.
                textBodyElement.Load(this.Shape.TextBody);
            }

            this.Update();

            return this;
        }
    }
}
