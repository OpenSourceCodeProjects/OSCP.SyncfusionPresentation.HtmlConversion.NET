using Syncfusion.Presentation;
using OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    /// <summary>
    /// Convert an HTML presentation document to a Syncfusion.Presentaton.IPresentation.
    /// </summary>
    public class HtmlToPresentationConverter
    {
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
            PptxDocument pptxDocument = new PptxDocument();
            pptxDocument.Load(xml);

            // Get the IPresentation created from the HTML.
            return pptxDocument.IPresentation;
        }
    }
}
