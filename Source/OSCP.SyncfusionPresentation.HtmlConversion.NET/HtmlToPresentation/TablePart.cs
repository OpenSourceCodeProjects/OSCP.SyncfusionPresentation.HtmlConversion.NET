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
    internal class TablePart : PartObject
    {
        internal SlidePart SlidePart { set; get; }
        internal ITable ITable { get; set; }
        internal double Left { get; private set; }
        internal double Top { get; private set; }
        internal double Width { get; private set; }
        internal double Height { get; private set; }

        internal TablePart(SlidePart slidePart)
        {
            this.SlidePart = slidePart;
        }

        internal virtual void Load(XmlNode tableNode)
        {
            this.Node = tableNode;

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

            // Get the built-in table style.
            string builtinStyleAttributeValue = this.Attr("data-builtin-style");
            if (builtinStyleAttributeValue.Length > 0)
            {
                BuiltInTableStyle builtInTableStyle;

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

                // Populate the table row.
                for (cdx = 0; cdx < tableCells.Count; cdx++)
                {
                    TableCellPart tableCellPart = new TableCellPart(this);
                    tableCellPart.Load(tableCells[cdx], rdx + 1, cdx);
                }
            }
        }
    }
}
