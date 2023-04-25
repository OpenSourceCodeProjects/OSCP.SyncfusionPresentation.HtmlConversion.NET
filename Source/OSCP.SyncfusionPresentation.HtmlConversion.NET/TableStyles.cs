using Syncfusion.Drawing;
using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal struct TableLineStyle
    {
        internal Color Color;
        internal int Width;
        internal LineStyle Style;

        internal string ToCss()
        {
            string lineStyle = "solid";
            return $"{this.Width}px {lineStyle} {this.Color.ToCss()}";
        }

        internal static TableLineStyle Empty
        {
            get
            {
                return new TableLineStyle
                {
                    Color = Color.Empty,
                    Width = 0
                };
            }
        }
    }

    internal struct TableBorderStyle
    {
        internal TableLineStyle Top;
        internal TableLineStyle Right;
        internal TableLineStyle Bottom;
        internal TableLineStyle Left;
    }

    internal struct TableStyle
    {
        internal Color HeadingFillColor;
        internal Color FillColor;
        internal Color BandedFillColor;
        internal TableBorderStyle TableBorder;
        internal TableBorderStyle HeaderBorder;
        internal TableBorderStyle RowBorder;
        internal bool BandedRows;
        internal bool BandedColumns;
    }

    internal struct TableStyles
    {
        internal static Dictionary<BuiltInTableStyle, TableStyle> List = new Dictionary<BuiltInTableStyle, TableStyle>
        {
            #region No Accent (Black and White)

            // Themed Style
            {
                BuiltInTableStyle.NoStyleTableGrid, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 255, 255, 255),
                    FillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },

            // Light Style
            {
                BuiltInTableStyle.LightStyle1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 0, 0, 0),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 0, 0, 0),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Medium Style
            {
                BuiltInTableStyle.MediumStyle1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 231, 231, 231),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 203, 203, 203),
                    BandedFillColor = Color.FromArgb(0, 231, 231, 231),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 231, 231, 231),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 231, 231, 231),
                    FillColor = Color.FromArgb(0, 203, 203, 203),
                    BandedFillColor = Color.FromArgb(0, 231, 231, 231),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Dark Style
            {
                BuiltInTableStyle.DarkStyle1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 203, 203, 203),
                    BandedFillColor = Color.FromArgb(0, 231, 231, 231),
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.DarkStyle2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 203, 203, 203),
                    BandedFillColor = Color.FromArgb(0, 231, 231, 231),
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            #endregion No Accent (Black and White)

            
            #region Accent 1 (Blue)

            // Themed Style
            {
                BuiltInTableStyle.ThemedStyle1Accent1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 68, 114, 196),
                    FillColor = Color.FromArgb(0, 123, 152, 210),
                    BandedFillColor = Color.FromArgb(0, 154, 171, 217),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 124, 152, 210), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.ThemedStyle2Accent1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 92, 129, 203),
                    FillColor = Color.FromArgb(0, 112, 147, 213),
                    BandedFillColor = Color.FromArgb(0, 64, 113, 202),
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Light Style
            {
                BuiltInTableStyle.LightStyle1Accent1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 68, 114, 196),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle2Accent1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 68, 114, 196),
                    FillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle3Accent1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 68, 114, 196),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Medium Style
            {
                BuiltInTableStyle.MediumStyle1Accent1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 68, 114, 196),
                    FillColor = Color.FromArgb(0, 233, 235, 245),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 88, 129, 202), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 88, 129, 202), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 88, 129, 202), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 88, 129, 202), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 88, 129, 202), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle2Accent1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 68, 114, 196),
                    FillColor = Color.FromArgb(0, 207, 213, 234),
                    BandedFillColor = Color.FromArgb(0, 233, 235, 245),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle3Accent1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 68, 114, 196),
                    FillColor = Color.FromArgb(0, 231, 231, 231),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle4Accent1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 233, 235, 245),
                    FillColor = Color.FromArgb(0, 207, 213, 234),
                    BandedFillColor = Color.FromArgb(0, 233, 235, 245),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 68, 114, 196), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Dark Style
            {
                BuiltInTableStyle.DarkStyle1Accent1, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 52, 89, 156),
                    BandedFillColor = Color.FromArgb(0, 68, 114, 196),
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            #endregion Accent 1 (Blue)

            
            #region Accent 2 (Orange)

            // Themed Style
            {
                BuiltInTableStyle.ThemedStyle1Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 237, 125, 49),
                    FillColor = Color.FromArgb(0, 243, 160, 114),
                    BandedFillColor = Color.FromArgb(0, 245, 178, 151),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 237, 128, 56), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 237, 128, 56), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 237, 128, 56), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 237, 128, 56), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.ThemedStyle2Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 241, 139, 81),
                    FillColor = Color.FromArgb(0, 245, 157, 101),
                    BandedFillColor = Color.FromArgb(0, 245, 125, 45),
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Light Style
            {
                BuiltInTableStyle.LightStyle1Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 237, 125, 49),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 2, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle2Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 237, 125, 49),
                    FillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle3Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 237, 125, 49),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Medium Style
            {
                BuiltInTableStyle.MediumStyle1Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 237, 125, 49),
                    FillColor = Color.FromArgb(0, 252, 236, 232),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle2Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 237, 125, 49),
                    FillColor = Color.FromArgb(0, 248, 215, 205),
                    BandedFillColor = Color.FromArgb(0, 252, 236, 232),
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle3Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 237, 125, 49),
                    FillColor = Color.FromArgb(0, 231, 231, 231),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle4Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 252, 236, 232),
                    FillColor = Color.FromArgb(0, 248, 215, 205),
                    BandedFillColor = Color.FromArgb(0, 252, 236, 232),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 237, 125, 49), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Dark Style
            {
                BuiltInTableStyle.DarkStyle1Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 189, 98, 37),
                    BandedFillColor = Color.FromArgb(0, 237, 125, 49),
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            #endregion Accent 2 (Orange)

            
            #region Accent 3 (Gray)

            // Themed Style
            {
                BuiltInTableStyle.ThemedStyle1Accent3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 165, 165, 165),
                    FillColor = Color.FromArgb(0, 189, 189, 189),
                    BandedFillColor = Color.FromArgb(0, 201, 201, 201),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 169, 169, 169), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 169, 169, 169), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 169, 169, 169), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 169, 169, 169), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.ThemedStyle2Accent3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 174, 174, 174),
                    FillColor = Color.FromArgb(0, 187, 187, 187),
                    BandedFillColor = Color.FromArgb(0, 166, 166, 166),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 223, 223, 223), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 223, 223, 223), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 223, 223, 223), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 223, 223, 223), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Light Style
            {
                BuiltInTableStyle.LightStyle1Accent3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 165, 165, 165),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle2Accent3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 165, 165, 165),
                    FillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle3Accent3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 165, 165, 165),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Medium Style
            {
                BuiltInTableStyle.MediumStyle1Accent3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 165, 165, 165),
                    FillColor = Color.FromArgb(0, 240, 240, 240),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle2Accent3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 165, 165, 165),
                    FillColor = Color.FromArgb(0, 225, 225, 225),
                    BandedFillColor = Color.FromArgb(0, 240, 240, 240),
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle3Accent3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 165, 165, 165),
                    FillColor = Color.FromArgb(0, 231, 231, 231),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle4Accent3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 240, 240, 240),
                    FillColor = Color.FromArgb(0, 225, 225, 225),
                    BandedFillColor = Color.FromArgb(0, 240, 240, 240),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 165, 165, 165), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Dark Style
            {
                BuiltInTableStyle.DarkStyle1Accent3, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 131, 131, 131),
                    BandedFillColor = Color.FromArgb(0, 165, 165, 165),
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            #endregion Accent 3 (Gray)


            #region Accent 4 (Yellow)

            // Themed Style
            {
                BuiltInTableStyle.ThemedStyle1Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 255, 192, 0),
                    FillColor = Color.FromArgb(0, 255, 208, 89),
                    BandedFillColor = Color.FromArgb(0, 255, 216, 144),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 195, 18), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 195, 18), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 195, 18), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 195, 18), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.ThemedStyle2Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 255, 199, 60),
                    FillColor = Color.FromArgb(0, 255, 209, 79),
                    BandedFillColor = Color.FromArgb(0, 255, 198, 9),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 219, 155), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 219, 155), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 219, 155), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 219, 155), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Light Style
            {
                BuiltInTableStyle.LightStyle1Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 255, 192, 0),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle2Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 255, 192, 0),
                    FillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle3Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 255, 192, 0),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Medium Style
            {
                BuiltInTableStyle.MediumStyle1Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 255, 192, 0),
                    FillColor = Color.FromArgb(0, 255, 244, 231),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle2Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 255, 192, 0),
                    FillColor = Color.FromArgb(0, 255, 232, 203),
                    BandedFillColor = Color.FromArgb(0, 255, 244, 231),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle3Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 255, 192, 0),
                    FillColor = Color.FromArgb(0, 231, 231, 231),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle4Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 255, 244, 231),
                    FillColor = Color.FromArgb(0, 255, 232, 203),
                    BandedFillColor = Color.FromArgb(0, 255, 244, 231),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 192, 0), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Dark Style
            {
                BuiltInTableStyle.DarkStyle1Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 203, 153, 0),
                    BandedFillColor = Color.FromArgb(0, 255, 192, 0),
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            #endregion Accent 4 (Yellow)


            #region Accent 5 (Light Blue)

            // Themed Style
            {
                BuiltInTableStyle.ThemedStyle1Accent5, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 91, 155, 213),
                    FillColor = Color.FromArgb(0, 137, 180, 224),
                    BandedFillColor = Color.FromArgb(0, 163, 193, 229),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.ThemedStyle2Accent5, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 91, 155, 213),
                    FillColor = Color.FromArgb(0, 137, 180, 224),
                    BandedFillColor = Color.FromArgb(0, 88, 156, 219),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 219, 155), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 219, 155), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 219, 155), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 219, 155), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Light Style
            {
                BuiltInTableStyle.LightStyle1Accent5, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 91, 155, 213),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle2Accent5, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 91, 155, 213),
                    FillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle3Accent5, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 91, 155, 213),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Medium Style
            {
                BuiltInTableStyle.MediumStyle1Accent5, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 91, 155, 213),
                    FillColor = Color.FromArgb(0, 234, 239, 247),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle2Accent5, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 91, 155, 213),
                    FillColor = Color.FromArgb(0, 210, 222, 239),
                    BandedFillColor = Color.FromArgb(0, 234, 239, 247),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle3Accent5, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 91, 155, 213),
                    FillColor = Color.FromArgb(0, 231, 231, 231),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle4Accent5, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 234, 239, 247),
                    FillColor = Color.FromArgb(0, 210, 222, 239),
                    BandedFillColor = Color.FromArgb(0, 234, 239, 247),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 91, 155, 213), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Dark Style
            {
                BuiltInTableStyle.DarkStyle1Accent5, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 71, 123, 169),
                    BandedFillColor = Color.FromArgb(0, 91, 155, 213),
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            #endregion Accent 5 (Light Blue)


            #region Accent 6 (Green)

            // Themed Style
            {
                BuiltInTableStyle.ThemedStyle1Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 112, 173, 71),
                    FillColor = Color.FromArgb(0, 150, 194, 123),
                    BandedFillColor = Color.FromArgb(0, 170, 206, 153),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.ThemedStyle2Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 122, 182, 85),
                    FillColor = Color.FromArgb(0, 145, 195, 113),
                    BandedFillColor = Color.FromArgb(0, 112, 173, 71),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Light Style
            {
                BuiltInTableStyle.LightStyle1Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 112, 173, 71),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle2Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 112, 173, 71),
                    FillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle3Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(1, 255, 255, 255),
                    FillColor = Color.FromArgb(204, 112, 173, 71),
                    BandedFillColor = Color.FromArgb(1, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Medium Style
            {
                BuiltInTableStyle.MediumStyle1Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 112, 173, 71),
                    FillColor = Color.FromArgb(0, 235, 241, 233),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle2Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 112, 173, 71),
                    FillColor = Color.FromArgb(0, 213, 227, 207),
                    BandedFillColor = Color.FromArgb(0, 235, 241, 233),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle3Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 112, 173, 71),
                    FillColor = Color.FromArgb(0, 231, 231, 231),
                    BandedFillColor = Color.FromArgb(0, 255, 255, 255),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 0, 0, 0), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle4Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 235, 241, 233),
                    FillColor = Color.FromArgb(0, 213, 227, 207),
                    BandedFillColor = Color.FromArgb(0, 235, 241, 233),
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = Color.FromArgb(0, 112, 173, 71), Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Dark Style
            {
                BuiltInTableStyle.DarkStyle1Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 0, 0, 0),
                    FillColor = Color.FromArgb(0, 88, 137, 55),
                    BandedFillColor = Color.FromArgb(0, 112, 173, 71),
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = Color.FromArgb(0, 255, 255, 255), Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            #endregion Accent 6 (Green)


            #region Dark Style Two Tone

            {
                BuiltInTableStyle.DarkStyle2Accent1Accent2, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 237, 125, 49),
                    FillColor = Color.FromArgb(0, 207, 213, 234),
                    BandedFillColor = Color.FromArgb(0, 233, 235, 245),
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.DarkStyle2Accent3Accent4, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 255, 192, 0),
                    FillColor = Color.FromArgb(0, 225, 225, 225),
                    BandedFillColor = Color.FromArgb(0, 240, 240, 240),
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.DarkStyle2Accent5Accent6, new TableStyle
                {
                    HeadingFillColor = Color.FromArgb(0, 112, 173, 71),
                    FillColor = Color.FromArgb(0, 210, 222, 239),
                    BandedFillColor = Color.FromArgb(0, 234, 239, 247),
                    BandedRows = true,
                    BandedColumns = false
                }
            }

            #endregion Dark Style Two Tone
        };
    }
}
