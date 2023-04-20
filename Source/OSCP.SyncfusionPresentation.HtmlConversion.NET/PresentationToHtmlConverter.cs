using Syncfusion.Presentation;
using OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    public class PresentationToHtmlConverter
    {
        private HtmlDocument HtmlDocument { get; set; }

        public PresentationToHtmlConverter()
        {
        }

        public string Convert(IPresentation presentation)
        {
            this.HtmlDocument = new HtmlDocument();
            this.HtmlDocument.Load(presentation);

            return this.HtmlDocument.OuterHtml;
        }
    }
}
