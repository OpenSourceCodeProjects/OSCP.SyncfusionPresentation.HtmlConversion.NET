using OSCP.SyncfusionPresentation.HtmlConversion.NET.PresentationToHtml;
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
    /// Convert an HTMLTableCellElement to a Syncfusion.Presentaton.ICell object.
    /// </summary>
    internal class TableCellPart : PartObject, IPartWithTextBody
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

        /// <summary>
        /// Syncfusion.Presentaton.ICell object that is created.
        /// </summary>
        internal ICell ICell { get; set; }

        /// <summary>
        /// Parent TablePart object.
        /// </summary>
        internal TablePart TablePart { get; set; }

        /// <summary>
        /// Row for the cell.
        /// </summary>
        internal int Row { get; set; }

        /// <summary>
        /// Column for the cell.
        /// </summary>
        internal int Col { get; set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="tablePart">Parent TablePart object.</param>
        internal TableCellPart(TablePart tablePart)
        {
            this.TablePart = tablePart;
        }

        /// <summary>
        /// Load the TableCellPart.
        /// </summary>
        /// <param name="tableNode">XmlNode that represents the HTMLTableCellElement.</param>
        /// <param name="row">Row for the cell.</param>
        /// <param name="col">Column for the cell.</param>
        internal void Load(XmlNode tableCellNode, int row, int col)
        {
            // Store parameter values passed in.
            this.Node = tableCellNode;
            this.Row = row;
            this.Col = col;

            // Add the Syncfusion.Presentation.ICell.
            this.ICell = this.TablePart.ITable[row, col];

            // Apply borders.
            this.ApplyBorder("top", this.Css("border-top"));
            this.ApplyBorder("right", this.Css("border-right"));
            this.ApplyBorder("bottom", this.Css("border-bottom"));
            this.ApplyBorder("left", this.Css("border-left"));

            // Get the background color for the cell if it has one.
            string backgroundColor = this.Css("background-color");

            // The cell has a background color.
            if (backgroundColor.Length > 0) 
            {
                // Apply background color.
                this.ICell.Fill.FillType = FillType.Solid;
                this.ICell.Fill.SolidFill.Color = ColorExtensions.FromCss(backgroundColor);
            }

            // Get the child div element.
            XmlNode childNode = tableCellNode.SelectSingleNode("div");

            // There is a child element and it has a TextBody CSS class.
            if (childNode != null && childNode.Attributes["class"] != null && childNode.Attributes["class"].Value.IndexOf(PptxDocument.Settings.CssClass.TextBody) > -1)
            {
                // Add a text body.
                TextBodyPart textBodyPart = new TextBodyPart(this);
                // Load the text body.
                textBodyPart.Load(childNode);
            }
        }

        /// <summary>
        /// Apply a border to the ICell.
        /// </summary>
        /// <param name="borderPosition">Border position: top, right, bottom or left.</param>
        /// <param name="borderStyleValue">Border CSS style value.</param>
        private void ApplyBorder(string borderPosition, string borderStyleValue)
        {
            // A border style value was passed in.
            if (borderStyleValue.Length > 0) 
            {
                // Get the appropriate ILineFormat from the ICell based on the border position passed in.
                ILineFormat lineFormat = borderPosition == "top"
                    ? this.ICell.CellBorders.BorderTop
                    : (
                        borderPosition == "right"
                            ? this.ICell.CellBorders.BorderRight
                            : (
                                borderPosition == "bottom"
                                    ? this.ICell.CellBorders.BorderBottom
                                    : this.ICell.CellBorders.BorderLeft
                            )
                    );

                // Split the border style value into the width, style and color.
                string[] borderStyleValues = borderStyleValue.Split(' ');

                // All 3 values found.
                if (borderStyleValues.Length == 3)
                {
                    // Set the ILineFormat values.
                    lineFormat.Weight = double.Parse(Regex.Match(borderStyleValues[0], @"[\.0-9]*").Value);
                    lineFormat.Style = LineStyle.Single;
                    lineFormat.Fill.FillType = FillType.Solid;
                    lineFormat.Fill.SolidFill.Color = ColorExtensions.FromCss(borderStyleValues[2]);
                }
            }
        }
    }
}
