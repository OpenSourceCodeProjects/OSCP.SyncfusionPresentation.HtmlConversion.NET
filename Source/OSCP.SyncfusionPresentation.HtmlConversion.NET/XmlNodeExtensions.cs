using OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal static class XmlNodeExtensions
    {
        internal static bool HasClass(this XmlNode xmlNode, string className)
        {
            bool hasClass = false;

            XmlAttribute attribute;

            if ((attribute = xmlNode.Attributes["class"]) != null)
            {
                hasClass = attribute.Value.IndexOf(className) > -1;
            }
            return hasClass;
        }

        internal static void Attr(this XmlNode xmlNode, string name, string value)
        {
            xmlNode.GetAttribute(name).Value = value;
        }

        internal static XmlAttribute GetAttribute(this XmlNode xmlNode, string name)
        {
            XmlAttribute attribute;

            if ((attribute = xmlNode.Attributes[name]) == null)
            {
                attribute = xmlNode.OwnerDocument.CreateAttribute(name);
                xmlNode.Attributes.Append(attribute);
            }
            return attribute;
        }
    }
}
