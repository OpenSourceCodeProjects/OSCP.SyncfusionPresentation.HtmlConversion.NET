using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    public class ShapeElement : HtmlElement
    {
        internal IShape Shape { get; set; }
        internal SlideElement Parent { get; set; }

        public ShapeElement(XmlNode node) : base(node)
        {
            this.AddClass("pptx-shape");
        }

        internal ShapeElement Load(IShape shape)
        {
            this.Shape = shape;

            this.Css("border", "1px dashed gray")
                .Css("position", "absolute")
                .Css("display", "table")
                .Css("top", $"{shape.Top}px")
                .Css("left", $"{shape.Left}px")
                .Css("height", $"{shape.Height}px")
                .Css("width", $"{shape.Width}px");

            List<UnorderedListElement> unorderedListElements = new List<UnorderedListElement>();

            // Loop over all the paragraphs in the shape.
            foreach (IParagraph paragraph in shape.TextBody.Paragraphs)
            {
                // Paragraph is text.
                if (paragraph.ListFormat.Type == ListType.None)
                {
                    // Create a paragraph element.
                    ParagraphElement paragraphElement = this.AddElement<ParagraphElement>("p");
                    paragraphElement.Parent = this;

                    // Load the paragraph element from the Syncfusion paragraph.
                    paragraphElement.Load(paragraph);
                }
                // Paragraph is a bullet list.
                else if (paragraph.ListFormat.Type == ListType.Bulleted)
                {
                    if (paragraph.IndentLevelNumber > (unorderedListElements.Count - 1))
                    {
                        UnorderedListElement unorderedListElement = null;

                        if (paragraph.IndentLevelNumber == 0)
                        {
                            // Create an unordered list element at the root of the shape element.
                            unorderedListElement = this.AddElement<UnorderedListElement>("ul");
                        }
                        else
                        {
                            unorderedListElement = unorderedListElements[paragraph.IndentLevelNumber - 1].AddElement<UnorderedListElement>("ul");
                        }
                        unorderedListElement.Parent = this;
                        unorderedListElement.Load(paragraph);

                        unorderedListElements.Add(unorderedListElement);
                    }

                    unorderedListElements[paragraph.IndentLevelNumber].AddListItem(paragraph.Text);
                }
                // Paragraph is a numbered list.
                else if (paragraph.ListFormat.Type == ListType.Numbered)
                {

                }
                // Paragraph is a picture.
                else if (paragraph.ListFormat.Type == ListType.Picture)
                {

                }
            }

            this.Update();

            return this;
        }
    }
}
