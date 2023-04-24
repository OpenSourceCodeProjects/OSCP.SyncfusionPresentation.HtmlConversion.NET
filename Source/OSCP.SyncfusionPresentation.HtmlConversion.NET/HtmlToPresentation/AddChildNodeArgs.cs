using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    delegate void AddChildNodeDelegate(AddChildNodeArgs e);

    internal class AddChildNodeArgs
    {
        internal XmlNode XmlNode { get; set; }
    }
}
