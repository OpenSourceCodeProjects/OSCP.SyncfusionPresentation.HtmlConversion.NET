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
    internal class SlideItemPart : PartObject
    {
        internal SlidePart Parent { get; set; }

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
        internal SlideItemPart(SlidePart slidePart)
        {
            this.Parent = slidePart;
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
        }
    }
}
