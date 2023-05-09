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
        public string LayoutSlides { get; set; }
        public string LayoutSlide { get; set; }
        public string Slides { get; set; }
        public string Slide { get; set; }
        public string SlideItem { get; set; }
        public string AutoShape { get; set; }
        public string Placeholder { get; set; }
        public string Picture { get; set; }
        public string Table { get; set; }
        public string TextBody { get; set; }
        public string Paragraph { get; set; }
        public string OrderedList { get; set; }
        public string UnorderedList { get; set; }
        public string TextPart { get; set; }

        internal CssClassSettings()
        {
            this.Document = "pptx-document";
            this.LayoutSlides = "pptx-layout-slides";
            this.LayoutSlide = "pptx-layout-slide";
            this.Slides = "pptx-slides";
            this.Slide = "pptx-slide";
            this.SlideItem = "pptx-slide-item";
            this.AutoShape = "pptx-autoshape";
            this.Placeholder = "pptx-placeholder";
            this.Picture = "pptx-picture";
            this.Table = "pptx-table";
            this.TextBody = "pptx-textbody";
            this.Paragraph = "pptx-paragraph";
            this.OrderedList = "pptx-ordered-list";
            this.UnorderedList = "pptx-unordered-list";
            this.TextPart = "pptx-textpart";
        }
    }

    public class ElementAttributes
    {
        public Dictionary<string, string> Document { get; set; }
        public Dictionary<string, string> LayoutSlides { get; set; }
        public Dictionary<string, string> LayoutSlide { get; set; }
        public Dictionary<string, string> Slides { get; set; }
        public Dictionary<string, string> Slide { get; set; }
        public Dictionary<string, string> SlideItem { get; set; }
        public Dictionary<string, string> AutoShape { get; set; }
        public Dictionary<string, string> Placeholder { get; set; }
        public Dictionary<string, string> Picture { get; set; }
        public Dictionary<string, string> Table { get; set; }
        public Dictionary<string, string> TextBody { get; set; }
        public Dictionary<string, string> Paragraph { get; set; }
        public Dictionary<string, string> OrderedList { get; set; }
        public Dictionary<string, string> UnorderedList { get; set; }
        public Dictionary<string, string> TextPart { get; set; }

        internal ElementAttributes()
        {
            this.Document = new Dictionary<string, string>();
            this.LayoutSlides = new Dictionary<string, string>();
            this.LayoutSlide = new Dictionary<string, string>();
            this.Slides = new Dictionary<string, string>();
            this.Slide = new Dictionary<string, string>();
            this.SlideItem = new Dictionary<string, string>();
            this.AutoShape = new Dictionary<string, string>();
            this.Placeholder = new Dictionary<string, string>();
            this.Picture = new Dictionary<string, string>();
            this.Table = new Dictionary<string, string>();
            this.TextBody = new Dictionary<string, string>();
            this.Paragraph = new Dictionary<string, string>();
            this.OrderedList = new Dictionary<string, string>();
            this.UnorderedList = new Dictionary<string, string>();
            this.TextPart = new Dictionary<string, string>();
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

    public class PresentationToHtmlConverterSettings
    {
        public CssClassSettings CssClass { get; set; }
        public ElementAttributes ElementAttributes { get; set; }

        internal ImageDataSettings ImageData { get; set; }

        internal PresentationToHtmlConverterSettings()
        {
            this.CssClass = new CssClassSettings();
            this.ElementAttributes = new ElementAttributes();
            this.ImageData = new ImageDataSettings();
        }
    }

    public class HtmlToPresentationConverterSettings
    {
        public CssClassSettings CssClass { get; set; }

        internal ImageDataSettings ImageData { get; set; }

        internal HtmlToPresentationConverterSettings()
        {
            this.CssClass = new CssClassSettings();
            this.ImageData = new ImageDataSettings();
        }
    }
}
