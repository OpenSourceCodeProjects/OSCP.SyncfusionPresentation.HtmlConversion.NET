using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class TextBodyElement : HtmlElement
    {
        internal const string ELEMENT_NAME = "div";

        internal ITextBody ITextBody { get; set; }
        internal HtmlElement Parent { get; set; }

        public TextBodyElement(XmlNode node) : base(node)
        {
            this.AddClass(HtmlDocument.Settings.CssClass.TextBody);
        }

        internal TextBodyElement Load(ITextBody textBody)
        {
            this.ITextBody = textBody;

            string verticalAlignment = textBody.VerticalAlignment.ToString().ToLower();

            this.Css("display", "table-cell")
                .Css("vertical-align", verticalAlignment == "none" ? "initial" : verticalAlignment);

            List<ListElement> listElements = new List<ListElement>();

            // Loop over all the paragraphs in the shape.
            foreach (IParagraph paragraph in this.ITextBody.Paragraphs)
            {
                // Paragraph is text.
                if (paragraph.ListFormat.Type == ListType.None || paragraph.ListFormat.Type == ListType.NotDefined)
                {
                    // Create a paragraph element.
                    ParagraphElement paragraphElement = this.AppendElement<ParagraphElement>(ParagraphElement.ELEMENT_NAME);
                    paragraphElement.Parent = this;

                    // Load the paragraph element from the Syncfusion paragraph.
                    paragraphElement.Load(paragraph);
                }
                // Paragraph is a bullet list.
                else if (paragraph.ListFormat.Type == ListType.Bulleted)
                {
                    // There isn't a list for the indent level.
                    if (paragraph.IndentLevelNumber > (listElements.Count - 1))
                    {
                        UnorderedListElement unorderedListElement = null;

                        // The list is at the top level.
                        if (paragraph.IndentLevelNumber == 0)
                        {
                            // Create an unordered list element at the root of the shape element.
                            unorderedListElement = this.AppendElement<UnorderedListElement>(UnorderedListElement.ELEMENT_NAME);
                        }
                        // The list is a "child" list.
                        else
                        {
                            // Add the unordered list element to the last list item at the indent level.
                            unorderedListElement = listElements[paragraph.IndentLevelNumber - 1].AddListToListItem<UnorderedListElement>();
                        }
                        // Set the shape as the parent of the list.
                        unorderedListElement.Parent = this;
                        // Load the list from the paragraph.
                        unorderedListElement.Load(paragraph);
                        // Store the unordered list.
                        listElements.Add(unorderedListElement);
                    }
                    // The indent level is lower then the last list that was added.
                    else if (paragraph.IndentLevelNumber < (listElements.Count - 1))
                    {
                        // Truncate the lists down to the indent level.
                        listElements = listElements.GetRange(0, (paragraph.IndentLevelNumber + 1));
                    }

                    // Add the list item.
                    listElements[paragraph.IndentLevelNumber].AddListItem(paragraph);
                }
                // Paragraph is a numbered list.
                else if (paragraph.ListFormat.Type == ListType.Numbered)
                {
                    // There isn't a list for the indent level.
                    if (paragraph.IndentLevelNumber > (listElements.Count - 1))
                    {
                        OrderedListElement orderedListElement = null;

                        // The list is at the top level.
                        if (paragraph.IndentLevelNumber == 0)
                        {
                            // Create an ordered list element at the root of the shape element.
                            orderedListElement = this.AppendElement<OrderedListElement>(OrderedListElement.ELEMENT_NAME);
                        }
                        // The list is a "child" list.
                        else
                        {
                            // Add the ordered list element to the last list item at the indent level.
                            orderedListElement = listElements[paragraph.IndentLevelNumber - 1].AddListToListItem<OrderedListElement>();
                        }
                        // Set the shape as the parent of the list.
                        orderedListElement.Parent = this;
                        // Load the list from the paragraph.
                        orderedListElement.Load(paragraph);
                        // Store the ordered list.
                        listElements.Add(orderedListElement);
                    }
                    // The indent level is lower then the last list that was added.
                    else if (paragraph.IndentLevelNumber < (listElements.Count - 1))
                    {
                        // Truncate the lists down to the indent level.
                        listElements = listElements.GetRange(0, (paragraph.IndentLevelNumber + 1));
                    }

                    // Add the list item.
                    listElements[paragraph.IndentLevelNumber].AddListItem(paragraph);
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
