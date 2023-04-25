using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    /// <summary>
    /// Base class for the OrderedListPart and UnorderedListPart.
    /// </summary>
    internal class ListPart : PartObject
    {
        internal int IndentLevel { get; set; }

        /// <summary>
        /// Parent ShapePart object.
        /// </summary>
        internal ShapePart ShapePart { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="shapePart">Parent ShapePart object.</param>
        internal ListPart(ShapePart shapePart)
        {
            this.ShapePart = shapePart;
        }
    }
}
