using Syncfusion.Presentation;
using OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    public class HtmlToPresentationConverter
    {
        public HtmlToPresentationConverter()
        {
        }

        public IPresentation Convert(string html)
        {
            // Cleanup HTML to be XML.
            string xml = html.Replace("&nbsp;", " ");

            PptxDocument pptxDocument = new PptxDocument();
            pptxDocument.Load(xml);

            return pptxDocument.PresentationDocument;
        }
    }
}
