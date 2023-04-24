using Syncfusion.Drawing;
using Syncfusion.Presentation;
using System.Collections.Generic;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml
{
    internal class TableElement : HtmlElement
    {
        internal const string ELEMENT_NAME = "table";

        internal SlideElement Parent { get; set; }
        internal ITable Table { get; set; }

        private HtmlElement THead { get; set; }
        private HtmlElement TBody { get; set; }
        private List<HtmlElement> Rows { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="node"></param>
        public TableElement(XmlNode node) : base(node)
        {
            // Add a CSS class to the element that identifies the presentation object.
            this.AddClass("pptx-table");

            // Initialize a list of row elements.
            this.Rows = new List<HtmlElement>();
        }

        /// <summary>
        /// Load the table element from the Syncfusion.Presentation.ITable object.
        /// </summary>
        /// <param name="table">Syncfusion.Presentation.ITable object</param>
        /// <returns></returns>
        internal TableElement Load(ITable table)
        {
            this.Table = table;

            // Set the attributes on the table element.
            this.Css("position", "absolute")
                .Css("display", "table")
                .Css("top", $"{table.Top}px")
                .Css("left", $"{table.Left}px")
                .Css("height", $"{table.Height}px")
                .Css("width", $"{table.Width}px")
                .Css("border-spacing", "0px")
                .Attr("data-builtin-style", table.BuiltInStyle.ToString());

            // Create a table body element.
            this.TBody = this.AppendElement<HtmlElement>("tbody");

            // Get from the built in table style that has been applied to the table.
            TableStyle tableStyle;
            TableStyles.List.TryGetValue(table.BuiltInStyle, out tableStyle);

            // Apply borders to the table.
            if (tableStyle.TableBorder.Top.Width > 0) this.Css($"border-top", tableStyle.TableBorder.Top.ToCss());
            if (tableStyle.TableBorder.Right.Width > 0) this.Css($"border-right", tableStyle.TableBorder.Right.ToCss());
            if (tableStyle.TableBorder.Bottom.Width > 0) this.Css($"border-bottom", tableStyle.TableBorder.Bottom.ToCss());
            if (tableStyle.TableBorder.Left.Width > 0) this.Css($"border-left", tableStyle.TableBorder.Left.ToCss());

            IRow row = null;
            ICell cell = null;
            bool isAlternateRow = false;
            bool isAlternateCell = false;
            Color fillColor;
            string cellElementName;

            // Loop over the rows in the table.
            for (int rdx = 0; rdx < table.Rows.Count; rdx++)
            {
                row = table.Rows[rdx];
                isAlternateRow = (rdx % 2) == 0;

                HtmlElement rowElement = null;

                if (rdx == 0 && table.HasHeaderRow == true)
                {
                    this.THead = this.PrependElement<HtmlElement>("thead");
                    rowElement = this.THead.AppendElement<HtmlElement>("tr");
                    cellElementName = "th";
                }
                else
                {
                    rowElement = this.TBody.AppendElement<HtmlElement>("tr");
                    cellElementName = "td";
                }

                for (int cdx = 0; cdx < row.Cells.Count; cdx++)
                {
                    cell = row.Cells[cdx];
                    isAlternateCell = (cdx % 2) == 1;

                    // Determine which style to get from the built in table styles.
                    fillColor = rdx == 0 
                        ? tableStyle.HeadingFillColor
                        : (
                            ((isAlternateRow == true && tableStyle.BandedRows == true) || (isAlternateCell == true && tableStyle.BandedColumns == true))
                                ? tableStyle.BandedFillColor
                                : tableStyle.FillColor
                        );

                    // Create a cell element.
                    HtmlElement cellElement = rowElement.AppendElement<HtmlElement>(cellElementName);

                    // Set the column and row span.
                    if (cell.ColumnSpan > 1) cellElement.Attr("colspan", cell.ColumnSpan.ToString());
                    if (cell.RowSpan > 1) cellElement.Attr("rowspan", cell.RowSpan.ToString());

                    // Apply a background color to the cell.
                    if (cell.Fill.FillType == FillType.Automatic)
                    {
                        cellElement.Css("background-color", fillColor.ToCss());
                    }
                    else if (cell.Fill.FillType == FillType.Solid)
                    {
                        cellElement.Css("background-color", cell.Fill.SolidFill.Color.SystemColor.ToCss());
                    }

                    // Not the heading row.
                    if (rdx == 0)
                    {
                        this.ApplyBorder(cellElement, "top", cell.CellBorders.BorderTop, tableStyle.HeaderBorder.Top);
                        this.ApplyBorder(cellElement, "right", cell.CellBorders.BorderRight, tableStyle.HeaderBorder.Right);
                        this.ApplyBorder(cellElement, "bottom", cell.CellBorders.BorderBottom, tableStyle.HeaderBorder.Bottom);
                        this.ApplyBorder(cellElement, "left", cell.CellBorders.BorderLeft, tableStyle.HeaderBorder.Left);
                    }
                    // Not the heading row.
                    else
                    {
                        this.ApplyBorder(cellElement, "top", cell.CellBorders.BorderTop, tableStyle.RowBorder.Top);
                        this.ApplyBorder(cellElement, "right", cell.CellBorders.BorderRight, tableStyle.RowBorder.Right);
                        this.ApplyBorder(cellElement, "bottom", cell.CellBorders.BorderBottom, tableStyle.RowBorder.Bottom);
                        this.ApplyBorder(cellElement, "left", cell.CellBorders.BorderLeft, tableStyle.RowBorder.Left);
                    }

                    // Display the text in the cell.
                    if (cell.TextBody != null)
                    {
                        foreach (IParagraph paragraph in cell.TextBody.Paragraphs)
                        {
                            ParagraphElement paragraphElement = cellElement.AppendElement<ParagraphElement>(ParagraphElement.ELEMENT_NAME);
                            paragraphElement.Parent = cellElement;
                            paragraphElement.Load(paragraph, cell.TextBody);
                        }
                    }

                    // Update the cell element attributes.
                    cellElement.Update();
                }
            }


            // Add the colgroup for the table.
            HtmlElement colGroupElement = this.PrependElement<HtmlElement>("colgroup");

            // Add col elements to the colgroup.
            for (int col = 0; col < table.ColumnsCount; col++)
            {
                colGroupElement.AppendElement<HtmlElement>("col");
            }

            // Update the table element attributes.
            this.Update();

            return this;
        }

        private void ApplyBorder(HtmlElement cellElement, string borderPosition, ILineFormat cellBorder, TableLineStyle defaultBorder)
        {
            if (cellBorder.Weight > 0)
            {
                string lineStyle = "solid";
                cellElement.Css($"border-{borderPosition}", $"{cellBorder.Weight}px {lineStyle} {cellBorder.Fill.SolidFill.Color.SystemColor.ToCss()}");
            }
            else if (defaultBorder.Width > 0)
            {
                cellElement.Css($"border-{borderPosition}", defaultBorder.ToCss());
            }
        }
    }
}
