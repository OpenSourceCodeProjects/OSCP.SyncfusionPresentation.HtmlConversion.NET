using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    /// <summary>
    /// Interface for a PartObject that has a Syncfusion.Presentation object that contains
    /// an ITextBody object.
    /// </summary>
    internal interface IPartWithTextBody
    {
        /// <summary>
        /// Get the ITextBody object from the Syncfusion.Presentation object.
        /// </summary>
        ITextBody ITextBody { get; }
    }
}
