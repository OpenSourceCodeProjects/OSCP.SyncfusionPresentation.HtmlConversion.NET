using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    internal class PartObject
    {
        private XmlNode _xmlNode;
        private Dictionary<string, string> _css;

        protected XmlNode Node 
        { 
            get
            {
                return this._xmlNode;
            }
            set
            {
                this._xmlNode = value;

                this._css = new Dictionary<string, string>();

                if (this._xmlNode.Attributes["style"] != null)
                {
                    string styleAttribute = this._xmlNode.Attributes["style"].Value;
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
                            this._css.Add(style.Substring(0, colonIndex), style.Substring(colonIndex + 1));
                        }
                    }
                }
            }
        }

        internal string Css(string styleName)
        {
            string styleValue;
            this._css.TryGetValue(styleName, out styleValue);
            return styleValue == null ? string.Empty : styleValue;
        }

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
    }
}
