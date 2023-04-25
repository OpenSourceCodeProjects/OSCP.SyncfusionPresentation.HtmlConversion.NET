using Syncfusion.Presentation;
using OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml;
using System.Collections.Generic;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    public class PresentationToHtmlConverter
    {
        private HtmlDocument HtmlDocument { get; set; }

        public ConverterSettings Settings { get; set; }

        public PresentationToHtmlConverter()
        {
            this.Settings = new ConverterSettings();
        }

        public string Convert(IPresentation presentation)
        {
            this.HtmlDocument = new HtmlDocument(this.Settings);
            this.HtmlDocument.Load(presentation);

            return this.HtmlDocument.OuterHtml;
        }

        public string Convert(IPresentation presentation, out List<Base64Image> base64Images)
        {
            base64Images = new List<Base64Image>();

            this.Settings.ImageData.Location = ImageDataLocation.Base64;
            this.Settings.ImageData.Base64Images = base64Images;

            this.HtmlDocument = new HtmlDocument(this.Settings);
            this.HtmlDocument.Load(presentation);

            return this.HtmlDocument.OuterHtml;
        }
    }
}
