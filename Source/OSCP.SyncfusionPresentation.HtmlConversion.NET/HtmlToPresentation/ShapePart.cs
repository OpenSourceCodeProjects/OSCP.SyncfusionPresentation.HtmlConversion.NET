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

            // Loop over all the child elements in the shape node.
            foreach (XmlNode childNode in shapeNode.ChildNodes)
            {
                // Child element is a paragraph.
                if (childNode.Name == "p")
                {
                    this.InitIShapeAsTextBox();

                    // Add a paragraph.
                    ParagraphPart paragraphPart = new ParagraphPart(this);
                    // Load the paragraph.
                    paragraphPart.Load(childNode);
                }
                // Child element is an unordered list.
                else if (childNode.Name == "ul")
                {

                }
                // Child element is an ordered list.
                else if (childNode.Name == "ol")
                {
                    this.InitIShapeAsTextBox();

                    // Add an ordered list.
                    OrderedListPart orderedListPart = new OrderedListPart(this);
                    // Load the ordered list.
                    orderedListPart.Load(childNode, 1);
                }
            }
        }

        private void InitIShapeAsTextBox()
        {
            // The Syncfusion.Presentation.IShape has not been created yet.
            if (this.IShape == null)
            {
                // Create an instance of the Syncfusion.Presentation.IShape as a text box.
                this.IShape = this.SlidePart.ISlide.AddTextBox(this.Left, this.Top, this.Width, this.Height);
            }
        }
    }
}
