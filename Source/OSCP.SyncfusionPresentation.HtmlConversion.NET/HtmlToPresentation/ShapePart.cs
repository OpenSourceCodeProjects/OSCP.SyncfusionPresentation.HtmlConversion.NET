using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    internal class ShapePart : PartObject
    {
        internal SlidePart SlidePart { set; get; }
        internal IShape IShape { get; set; }
        internal double Left { get; private set; }
        internal double Top { get; private set; }
        internal double Width { get; private set; }
        internal double Height { get; private set; }

        internal ShapePart(SlidePart slidePart)
        {
            this.SlidePart = slidePart;
        }

        internal virtual void Load(XmlNode shapeNode)
        {
            this.Node = shapeNode;

            this.Left = double.Parse(Regex.Match(this.Css("left"), @"[\.0-9]*").Value);
            this.Top = double.Parse(Regex.Match(this.Css("top"), @"[\.0-9]*").Value);
            this.Width = double.Parse(Regex.Match(this.Css("width"), @"[\.0-9]*").Value);
            this.Height = double.Parse(Regex.Match(this.Css("height"), @"[\.0-9]*").Value);

            if (shapeNode.FirstChild.Name == "p")
            {
                //this.IShape = this.SlidePart.ISlide.Shapes.AddShape(AutoShapeType.Rectangle, 100, 100, 100, 100);
                this.IShape = this.SlidePart.ISlide.AddTextBox(this.Left, this.Top, this.Width, this.Height);
                ParagraphPart paragraphPart = new ParagraphPart(this);
                paragraphPart.Load(shapeNode.FirstChild);
            }
        }
    }
}
