using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    public class CssClassSettings
    {
        public string Document { get; set; }
        public string Slides { get; set; }
        public string Slide { get; set; }
        public string Shape { get; set; }
        public string Table { get; set; }
        public string Paragraph { get; set; }
        public string OrderedList { get; set; }
        public string UnorderedList { get; set; }
        public string TextPart { get; set; }

        internal CssClassSettings()
        {
            this.Document = "pptx-document";
            this.Slides = "pptx-slides";
            this.Slide = "pptx-slide";
            this.Shape = "pptx-shape";
            this.Table = "pptx-table";
            this.Paragraph = "pptx-paragraph";
            this.OrderedList = "pptx-ordered-list";
            this.UnorderedList = "pptx-unordered-list";
            this.TextPart = "pptx-textpart";
        }
    }

    public enum ImageDataLocation
    {
        Embedded = 1,
        Base64 = 2
        //File = 3
    }

    public class Base64Image
    {
        public string Id { get; set; }
        public string ElementName { get; set; }
        public string ElementCssClass { get; set; }
        public string ImageType { get; set; }
        public string Data { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }

    public class ImageDataSettings
    {
        public ImageDataLocation Location { get; internal set; }

        internal List<Base64Image> Base64Images { get; set; }

        internal ImageDataSettings()
        {
            this.Location = ImageDataLocation.Embedded;
        }
    }

    public class ConverterSettings
    {
        public CssClassSettings CssClass { get; set; }

        internal ImageDataSettings ImageData { get; set; }

        internal ConverterSettings()
        {
            this.CssClass = new CssClassSettings();
            this.ImageData = new ImageDataSettings();
        }
    }
}
