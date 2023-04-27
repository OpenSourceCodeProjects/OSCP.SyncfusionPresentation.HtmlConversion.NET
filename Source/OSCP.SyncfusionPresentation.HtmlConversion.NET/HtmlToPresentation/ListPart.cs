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
        /// <summary>
        /// Indent level.
        /// </summary>
        internal int IndentLevel { get; set; }

        /// <summary>
        /// Parent TextBodyPart object.
        /// </summary>
        internal TextBodyPart TextBodyPart { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="textBodyPart">Parent TextBodyPart object.</param>
        internal ListPart(TextBodyPart textBodyPart)
        {
            // Store the parent text body part.
            this.TextBodyPart = textBodyPart;
        }
    }
}
