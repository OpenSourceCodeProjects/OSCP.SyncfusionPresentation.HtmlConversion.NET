using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET.HtmlToPresentation
{
    internal class TableCellPart : PartObject, ITextBodyPart
    {
        /// <summary>
        /// Get the ITextBody object.
        /// </summary>
        public ITextBody ITextBody
        {
            get
            {
                return this.ICell.TextBody;
            }
        }

        internal ICell ICell { get; set; }
        internal TablePart TablePart { get; set; }
        internal int Row { get; set; }
        internal int Col { get; set; }

        internal TableCellPart(TablePart tablePart)
        {
            this.TablePart = tablePart;
        }

        internal void Load(XmlNode tableCellNode, int row, int col)
        {
            this.Node = tableCellNode;
            this.Row = row;
            this.Col = col;

            // Add the Syncfusion.Presentation.ICell.
            this.ICell = this.TablePart.ITable[row, col];

            // Get the paragraph within the table cell.
            XmlNode paragraphNode = this.Node.SelectSingleNode("p");

            if (paragraphNode != null) 
            {
                // Add the table cell content.
                ParagraphPart paragraphPart = new ParagraphPart(this);
                paragraphPart.Load(paragraphNode);
            }
        }
    }
}
