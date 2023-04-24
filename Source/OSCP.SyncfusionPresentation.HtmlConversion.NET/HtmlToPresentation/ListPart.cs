using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    internal class ListPart : PartObject
    {
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
