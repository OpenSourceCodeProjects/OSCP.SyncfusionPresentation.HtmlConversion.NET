using Syncfusion.Presentation;
using OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation;
using OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml;
using System.Collections.Generic;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    /// <summary>
    /// Convert an HTML presentation document to a Syncfusion.Presentaton.IPresentation.
    /// </summary>
    public class HtmlToPresentationConverter
    {
        private PptxDocument PptxDocument { get; set; }

        public HtmlToPresentationConverterSettings Settings { get; private set; }

        public HtmlToPresentationConverter()
        {
            this.Settings = new HtmlToPresentationConverterSettings();
        }

        /// <summary>
        /// Convert an HTML presentation document to a Syncfusion.Presentaton.IPresentation.
        /// </summary>
        /// <param name="html">HTML that contains the presentation document.</param>
        /// <returns>Syncfusion.Presentaton.IPresentation.</returns>
        public IPresentation Convert(string html)
        {
            // Cleanup HTML to be XML.
            string xml = html.Replace("&nbsp;", " ");

            // Create a instance of the PptxDocument and load the XML.
            this.PptxDocument = new PptxDocument(this.Settings);
            this.PptxDocument.Load(xml);

            // Get the IPresentation created from the HTML.
            return this.PptxDocument.IPresentation;
        }

        /// <summary>
        /// Convert an HTML presentation document to a Syncfusion.Presentaton.IPresentation.
        /// </summary>
        /// <param name="html">HTML that contains the presentation document.</param>
        /// <param name="base64Images">List of images encoded as base 64 strings.</param>
        /// <returns>Syncfusion.Presentaton.IPresentation.</returns>
        public IPresentation Convert(string html, List<Base64Image> base64Images)
        {
            // Cleanup HTML to be XML.
            string xml = html.Replace("&nbsp;", " ");

            // Create a instance of the PptxDocument and load the XML.
            this.PptxDocument = new PptxDocument(this.Settings);
            this.PptxDocument.Load(xml, base64Images);

            // Get the IPresentation created from the HTML.
            return this.PptxDocument.IPresentation;
        }
    }
}
