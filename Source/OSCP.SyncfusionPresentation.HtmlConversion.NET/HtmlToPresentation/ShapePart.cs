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
    /// <summary>
    /// Convert an HTMLDivElement to a Syncfusion.Presentaton.IShape object.
    /// </summary>
    internal class ShapePart : PartObject, IPartWithTextBody
    {
        /// <summary>
        /// Get the ITextBody object.
        /// </summary>
        public ITextBody ITextBody
        {
            get
            {
                return this.IShape.TextBody;
            }
        }

        /// <summary>
        /// Parent SlidePart.
        /// </summary>
        internal SlidePart SlidePart { set; get; }

        /// <summary>
        /// Syncfusion.Presentaton.IShape object.
        /// </summary>
        internal IShape IShape { get; set; }

        /// <summary>
        /// Left position for the shape.
        /// </summary>
        internal double Left { get; private set; }

        /// <summary>
        /// Top position for the shape.
        /// </summary>
        internal double Top { get; private set; }

        /// <summary>
        /// Width of the shape.
        /// </summary>
        internal double Width { get; private set; }

        /// <summary>
        /// Height of the shape.
        /// </summary>
        internal double Height { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="slidePart">Parent SlidePart.</param>
        internal ShapePart(SlidePart slidePart)
        {
            this.SlidePart = slidePart;
        }

        /// <summary>
        /// Load the shape.
        /// </summary>
        /// <param name="shapeNode">XmlNode that represents the HTMLDivElement.</param>
        internal virtual void Load(XmlNode shapeNode)
        {
            this.Node = shapeNode;

            // Get the position of shape.
            this.Left = double.Parse(Regex.Match(this.Css("left"), @"[\.0-9]*").Value);
            this.Top = double.Parse(Regex.Match(this.Css("top"), @"[\.0-9]*").Value);
            this.Width = double.Parse(Regex.Match(this.Css("width"), @"[\.0-9]*").Value);
            this.Height = double.Parse(Regex.Match(this.Css("height"), @"[\.0-9]*").Value);

            // Get the child div element.
            XmlNode childNode = shapeNode.SelectSingleNode("div");

            // There is a child element and it has a TextBody CSS class.
            if (childNode != null && childNode.Attributes["class"] != null && childNode.Attributes["class"].Value.IndexOf(PptxDocument.Settings.CssClass.TextBody) > -1)
            {
                // Initialize the shape as a text box.
                this.InitIShapeAsTextBox();

                // Add a text body.
                TextBodyPart textBodyPart = new TextBodyPart(this);
                // Load the text body.
                textBodyPart.Load(childNode);
            }
        }

        /// <summary>
        /// Initialize the shape as a text box.
        /// </summary>
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
