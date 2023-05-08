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
    internal class PlaceholderPart : AutoShapePart
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="slidePart">Parent SlidePart.</param>
        internal PlaceholderPart(SlidePart slidePart) : base(slidePart)
        {
        }
    }
}
