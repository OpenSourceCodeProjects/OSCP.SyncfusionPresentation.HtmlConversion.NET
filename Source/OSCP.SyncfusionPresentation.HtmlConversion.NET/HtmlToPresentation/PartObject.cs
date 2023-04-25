using OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    /// <summary>
    /// Base class to convert HTML to a part o a PowerPoint presentation.
    /// </summary>
    internal class PartObject
    {
        private XmlNode _xmlNode;
        private Dictionary<string, string> _css;

        /// <summary>
        /// XML Node that represents an HTML element.
        /// </summary>
        protected XmlNode Node 
        { 
            get
            {
                return this._xmlNode;
            }
            set
            {
                this._xmlNode = value;

                // Parse the HTML element styles.
                this._css = this.ConvertStyleAttributeToDictionary(this._xmlNode);
            }
        }

        /// <summary>
        /// Get a CSS style for the HTML element.
        /// </summary>
        /// <param name="styleName"></param>
        /// <returns></returns>
        internal string Css(string styleName)
        {
            return this.Css(this._css, styleName);
        }

        /// <summary>
        /// Get a CSS style value from a dictionary.
        /// </summary>
        /// <param name="css">Dictionary that contains the CSS styles.</param>
        /// <param name="styleName">CSS style name.</param>
        /// <returns></returns>
        internal string Css(Dictionary<string, string> css, string styleName)
        {
            string styleValue;
            css.TryGetValue(styleName, out styleValue);
            return styleValue == null ? string.Empty : styleValue;
        }

        /// <summary>
        /// Try and get a CSS style value where the value is a double.
        /// </summary>
        /// <param name="styleName">Name of the CSS style.</param>
        /// <param name="styleValue">Value of the CSS style.</param>
        /// <returns>The CSS style was found.</returns>
        protected bool TryGetCssValue(string styleName, out double styleValue)
        {
            bool found = false;

            styleValue = 0F;

            string css = this.Css(styleName);

            if (css.Length > 0)
            {
                try
                {
                    styleValue = double.Parse(Regex.Match(css, @"[\.0-9]*").Value);
                    found = true;
                }
                catch
                {
                }
            }

            return found;
        }

        /// <summary>
        /// Convert the HTML element CSS styles to a dictionary.
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        protected Dictionary<string, string> ConvertStyleAttributeToDictionary(XmlNode xmlNode)
        {
            Dictionary<string, string> css = new Dictionary<string, string>();

            if (xmlNode.Attributes["style"] != null)
            {
                string styleAttribute = xmlNode.Attributes["style"].Value;
                string[] styles = styleAttribute.Split(';');
                int colonIndex;

                if (styles.Length > 1)
                {
                    for (int idx = 1; idx < styles.Length; idx++)
                    {
                        if (styles[idx].StartsWith("base64") == true)
                        {
                            styles[idx - 1] += ";" + styles[idx];
                            styles[idx] = string.Empty;
                        }
                    }
                }

                foreach (string style in styles)
                {
                    if ((colonIndex = style.IndexOf(':')) > -1)
                    {
                        css.Add(style.Substring(0, colonIndex), style.Substring(colonIndex + 1));
                    }
                }
            }

            return css;
        }

        /// <summary>
        /// Get a attribute value from the HTML element.
        /// </summary>
        /// <param name="name">Name of the attribute.</param>
        /// <returns>Attribute value.</returns>
        internal string Attr(string name)
        {
            string value = string.Empty;
            XmlAttribute attribute;

            if ((attribute = this.Node.Attributes[name]) != null)
            {
                value = attribute.Value;
            }
            return value;
        }
    }
}
