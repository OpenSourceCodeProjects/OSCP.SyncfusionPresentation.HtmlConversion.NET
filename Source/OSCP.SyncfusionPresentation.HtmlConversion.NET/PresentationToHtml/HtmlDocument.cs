using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    /// <summary>
    /// Create the HTML document from a Syncfusion Presentation document.
    /// </summary>
    internal class HtmlDocument
    {
        internal static PresentationToHtmlConverterSettings Settings { get; private set; }

        private XmlDocument XmlDocument { get; set; }
        private HtmlElement DocumentElement { get; set; }
        private HtmlElement SlidesElement { get; set; }
        private HtmlElement LayoutSlidesElement { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="settings">Settings for converting the presentation to HTML.</param>
        internal HtmlDocument(PresentationToHtmlConverterSettings settings)
        {
            // Store the settings so that it can be accessed by any of the HTML Element classes.
            HtmlDocument.Settings = settings;

            // Create the XML Document that will contain the HTML to be created.
            this.XmlDocument = new XmlDocument();

            // Add the Document Element to the XML Document.
            this.DocumentElement = this.AddElement("div");
            this.DocumentElement.AddClass(settings.CssClass.Document)
                .ApplyAttributes(settings.ElementAttributes.Document)
                .Update();

            // Add the Slides Element to the Document Element.
            this.SlidesElement = this.DocumentElement.AppendElement<HtmlElement>("div");
            this.SlidesElement.AddClass(settings.CssClass.Slides)
                .ApplyAttributes(settings.ElementAttributes.Slides)
                .Update();
        }

        /// <summary>
        /// Load the Presentation into an HTML document.
        /// </summary>
        /// <param name="presentation">Presentation to be converted to HTML.</param>
        internal void Load(IPresentation presentation)
        {
            // Loop over all the slides.
            foreach (ISlide slide in presentation.Slides)
            {
                // Add each slide to the slides element.
                SlideElement slideElement = this.AddSlide();
                slideElement.Load(slide);
            }

            // There are master slides.
            if (presentation.Masters.Count > 0) 
            {
                // Create the element for the layout slides.
                this.LayoutSlidesElement = this.DocumentElement.AppendElement<HtmlElement>("div");
                this.LayoutSlidesElement.AddClass(HtmlDocument.Settings.CssClass.LayoutSlides)
                    .ApplyAttributes(HtmlDocument.Settings.ElementAttributes.LayoutSlides)
                    .Update();

                // Loop over the master layout slides.
                foreach (ILayoutSlide layoutSlide in presentation.Masters.First().LayoutSlides)
                {
                    // Add each layout slide to the layout slides element.
                    LayoutSlideElement layoutSlideElement = this.LayoutSlidesElement.AppendElement<LayoutSlideElement>(LayoutSlideElement.ELEMENT_NAME);
                    layoutSlideElement.Parent = this.LayoutSlidesElement;
                    layoutSlideElement.Load(layoutSlide);
                }
            }
        }

        /// <summary>
        /// Add a slide to the slides element.
        /// </summary>
        /// <returns>New slide added to the slides element.</returns>
        internal SlideElement AddSlide()
        {
            SlideElement slideElement = this.SlidesElement.AppendElement<SlideElement>(SlideElement.ELEMENT_NAME);
            slideElement.Parent = this.SlidesElement;
            return slideElement;
        }

        /// <summary>
        /// Add an HTML element to the XML Document.
        /// </summary>
        /// <param name="name">Name for the element.</param>
        /// <returns>HTML element added to the XML document.</returns>
        private HtmlElement AddElement(string name)
        {
            // Create the XML Node.
            XmlNode node = this.XmlDocument.CreateElement(name);
            // Set configuration settings for the XML Node.
            (node as XmlElement).IsEmpty = false;
            // Add the XML Node to the XML Document.
            this.XmlDocument.AppendChild(node);
            // Create a new HTML Element and pass in the XML Node.
            HtmlElement htmlElement = new HtmlElement(node);

            // Return the HTML Element.
            return htmlElement;
        }

        /// <summary>
        /// Get the HTML for the presentation.
        /// </summary>
        internal string OuterHtml
        {
            get
            {
                return this.XmlDocument.OuterXml;
            }
        }
    }
}
