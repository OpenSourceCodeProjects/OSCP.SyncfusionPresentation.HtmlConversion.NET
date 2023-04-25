using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    /// <summary>
    /// Define a delegate for the AddChildNode event.
    /// </summary>
    /// <param name="e">Contains the child XmlNode.</param>
    delegate void AddChildNodeDelegate(AddChildNodeArgs e);

    /// <summary>
    /// Argument passed to the AddChildNode event handler.
    /// </summary>
    internal class AddChildNodeArgs
    {
        internal XmlNode XmlNode { get; set; }
    }
}
