using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    public class HtmlElement
    {
        private XmlNode Node { get; set; }
        internal string Name 
        { 
            get
            {
                return this.Node.Name;
            }
        }
        private List<string> Classes { get; set; }
        private List<KeyValuePair<string, string>> Styles { get; set; }

        public HtmlElement(XmlNode node)
        {
            this.Node = node;

            this.Classes = new List<string>();
            this.Styles = new List<KeyValuePair<string, string>>();
        }

        internal T AppendElement<T>(string name)
        {
            XmlNode node = this.Node.OwnerDocument.CreateElement(name);
            (node as XmlElement).IsEmpty = false;
            this.Node.AppendChild(node);
            T htmlElement = (T)Activator.CreateInstance(typeof(T), new object[] { node });

            return htmlElement;
        }

        internal T PrependElement<T>(string name)
        {
            XmlNode node = this.Node.OwnerDocument.CreateElement(name);
            (node as XmlElement).IsEmpty = false;
            this.Node.PrependChild(node);
            T htmlElement = (T)Activator.CreateInstance(typeof(T), new object[] { node });

            return htmlElement;
        }

        internal HtmlElement AddClass(string[] classNames)
        {
            this.Classes.AddRange(classNames);

            this.UpdateClassAttribute();

            return this;
        }

        internal HtmlElement AddClass(string className)
        {
            this.Classes.Add(className);

            return this;
        }

        internal HtmlElement Css(string style, string value) 
        {
            bool found = false;

            // Check whether the style exists already.
            for (int idx = 0; idx < this.Styles.Count && found == false; idx++)
            {
                // The style exists.
                if ((found = this.Styles[idx].Key == style) == true)
                {
                    // Remove the style.
                    this.Styles.RemoveAt(idx);
                }
            }

            // Add the style.
            this.Styles.Add(new KeyValuePair<string, string>(style, value));

            return this;
        }

        internal HtmlElement Attr(string name, string value)
        {
            this.Node.GetAttribute(name).Value = value;

            return this;
        }

        internal HtmlElement Text(string text)
        {
            this.Node.InnerText = text;

            return this;
        }

        internal HtmlElement Html(string html)
        {
            this.Node.InnerXml = html;

            return this;
        }

        internal HtmlElement Update()
        {
            this.UpdateClassAttribute();

            this.UpdateStyleAttribute();

            return this;
        }

        protected void ApplyAttributes(Dictionary<string, string> attributes)
        {
            foreach (string attributeName in attributes.Keys)
            {
                this.Node.GetAttribute(attributeName).Value = attributes[attributeName];
            }
        }

        private void UpdateClassAttribute()
        {
            XmlAttribute classAttribute = this.Node.GetAttribute("class");

            string classes = string.Empty;

            foreach (string name in this.Classes)
            {
                classes += " " + name;
            }
            classes = classes.Trim();

            classAttribute.Value = classes;
        }

        private void UpdateStyleAttribute()
        {
            XmlAttribute styleAttribute = this.Node.GetAttribute("style");

            string styles = string.Empty;

            foreach (KeyValuePair<string, string> keyValuePair in this.Styles)
            {
                styles += $"{keyValuePair.Key}:{keyValuePair.Value};";
            }

            styleAttribute.Value = styles;
        }
    }
}
