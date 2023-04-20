using System.Collections.Generic;
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

                    this._css = new Dictionary<string, string>();
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
    }
}
