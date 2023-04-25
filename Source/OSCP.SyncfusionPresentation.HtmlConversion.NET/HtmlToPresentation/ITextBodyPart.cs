using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    /// <summary>
    /// Interface to be applied to a PartObject that is creating a Syncfusion.Presentation object that
    /// has an ITextBody.
    /// </summary>
    internal interface ITextBodyPart
    {
        /// <summary>
        /// Get the ITextBody object.
        /// </summary>
        ITextBody ITextBody { get; }
    }
}
