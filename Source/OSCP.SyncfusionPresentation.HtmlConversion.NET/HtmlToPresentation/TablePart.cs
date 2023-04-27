using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    /// <summary>
    /// Convert an HTMLTableElement to a Syncfusion.Presentaton.ITable object.
    /// </summary>
    internal class TablePart : PartObject
    {
        /// <summary>
        /// Parent SlidePart object.
        /// </summary>
        internal SlidePart SlidePart { set; get; }

        /// <summary>
        /// Syncfusion.Presentaton.ITable object that is created.
        /// </summary>
        internal ITable ITable { get; set; }

        /// <summary>
        /// Left position for the shape.
        /// </summary>
        internal double Left { get; private set; }

        /// <summary>
        /// Top position for the shape.
        /// </summary>
        internal double Top { get; private set; }

        /// <summary>
        /// Width of the shape.
        /// </summary>
        internal double Width { get; private set; }

        /// <summary>
        /// Height of the shape.
        /// </summary>
        internal double Height { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="slidePart">Parent SlidePart.</param>
        internal TablePart(SlidePart slidePart)
        {
            this.SlidePart = slidePart;
        }

        /// <summary>
        /// Load the TablePart.
        /// </summary>
        /// <param name="tableNode">XmlNode that represents the HTMLTableElement.</param>
        internal virtual void Load(XmlNode tableNode)
        {
            this.Node = tableNode;

            // Get the position of table.
            this.Left = double.Parse(Regex.Match(this.Css("left"), @"[\.0-9]*").Value);
            this.Top = double.Parse(Regex.Match(this.Css("top"), @"[\.0-9]*").Value);
            this.Width = double.Parse(Regex.Match(this.Css("width"), @"[\.0-9]*").Value);
            this.Height = double.Parse(Regex.Match(this.Css("height"), @"[\.0-9]*").Value);

            // Get the number of columns.
            int columnCount = this.Node.SelectNodes("colgroup/col").Count;
            XmlNode headerRow = this.Node.SelectSingleNode("thead/tr");
            XmlNodeList tableRows = this.Node.SelectNodes("tbody/tr");

            // Create an instance of a Syncfusion.Presentation.ITable.
            this.ITable = this.SlidePart.ISlide.Shapes.AddTable((tableRows.Count + (headerRow != null ? 1 : 0)), columnCount, this.Left, this.Top, this.Width, this.Height);

            this.ITable.HasHeaderRow = headerRow != null;

            // Get the built-in table style if one is defined for this table.
            string builtinStyleAttributeValue = this.Attr("data-builtin-style");

            BuiltInTableStyle builtInTableStyle = BuiltInTableStyle.None;

            // A built-in table style is defined for this table.
            if (builtinStyleAttributeValue.Length > 0)
            {
                // The enumeration for the built-in table style was found.
                if (Enum.TryParse(builtinStyleAttributeValue, out builtInTableStyle) == true)
                {
                    this.ITable.BuiltInStyle = builtInTableStyle;
                }
            }

            int cdx = 0;

            // There is a header row.
            if (this.ITable.HasHeaderRow == true)
            {
                XmlNodeList tableHeaders = headerRow.SelectNodes("th");

                // Populate the header row.
                for (cdx = 0; cdx < tableHeaders.Count; cdx++)
                {
                    TableCellPart tableCellPart = new TableCellPart(this);
                    tableCellPart.Load(tableHeaders[cdx], 0, cdx);
                }
            }

            // Loop over the table rows.
            for (int rdx = 0; rdx < tableRows.Count; rdx++) 
            {
                XmlNodeList tableCells = tableRows[rdx].SelectNodes("td");

                // Populate the table cells for the current row.
                for (cdx = 0; cdx < tableCells.Count; cdx++)
                {
                    TableCellPart tableCellPart = new TableCellPart(this);
                    tableCellPart.Load(tableCells[cdx], rdx + (this.ITable.HasHeaderRow == true ? 1 : 0), cdx);
                }
            }
        }
    }
}
