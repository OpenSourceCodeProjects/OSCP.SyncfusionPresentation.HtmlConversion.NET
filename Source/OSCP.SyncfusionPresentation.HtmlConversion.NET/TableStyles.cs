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
        internal ElementColor Color;
        internal int Width;
        internal LineStyle Style;
        internal string Css
        {
            get
            {
                string lineStyle = "solid";
                return $"{this.Width}px {lineStyle} {this.Color.Css}";
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
        internal ElementColor HeadingFillColor;
        internal ElementColor FillColor;
        internal ElementColor BandedFillColor;
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
            // No Accent (Black and White)
            {
                BuiltInTableStyle.NoStyleTableGrid, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    FillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle 
                    { 
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }, 
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }, 
                        Left = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single } 
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle 
                    { 
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    FillColor = new ElementColor { R = 0, G = 0, B = 0, A = 80 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle 
                    { 
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }, 
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single } 
                    },
                    RowBorder = new TableBorderStyle 
                    { 
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single } 
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 0, G = 0, B = 0, A = 0 },
                    FillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle3, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    FillColor = new ElementColor { R = 0, G = 0, B = 0, A = 80 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 0, G = 0, B = 0, A = 0 },
                    FillColor = new ElementColor { R = 231, G = 231, B = 231, A = 0 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 0, G = 0, B = 0, A = 0 },
                    FillColor = new ElementColor { R = 203, G = 203, B = 203, A = 0 },
                    BandedFillColor = new ElementColor { R = 231, G = 231, B = 231, A = 0 },
                    TableBorder = new TableBorderStyle 
                    { 
                        Top = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }, 
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single } 
                    },
                    HeaderBorder = new TableBorderStyle 
                    { 
                        Top = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }, 
                        Right = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single } 
                    },
                    RowBorder = new TableBorderStyle 
                    { 
                        Top = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }, 
                        Right = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single } 
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle3, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 0, G = 0, B = 0, A = 0 },
                    FillColor = new ElementColor { R = 231, G = 231, B = 231, A = 0 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle4, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 231, G = 231, B = 231, A = 0 },
                    FillColor = new ElementColor { R = 203, G = 203, B = 203, A = 0 },
                    BandedFillColor = new ElementColor { R = 231, G = 231, B = 231, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.DarkStyle1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 0, G = 0, B = 0, A = 0 },
                    FillColor = new ElementColor { R = 203, G = 203, B = 203, A = 0 },
                    BandedFillColor = new ElementColor { R = 231, G = 231, B = 231, A = 0 },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.DarkStyle2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 0, G = 0, B = 0, A = 0 },
                    FillColor = new ElementColor { R = 203, G = 203, B = 203, A = 0 },
                    BandedFillColor = new ElementColor { R = 231, G = 231, B = 231, A = 0 },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Accent 1 (Blue)
            {
                BuiltInTableStyle.ThemedStyle1Accent1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 68, G = 114, B = 196, A = 0 },
                    FillColor = new ElementColor { R = 123, G = 152, B = 210, A = 0 },
                    BandedFillColor = new ElementColor { R = 154, G = 171, B = 217, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 124, G = 152, B = 210, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.ThemedStyle2Accent1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 92, G = 129, B = 203, A = 0 },
                    FillColor = new ElementColor { R = 112, G = 147, B = 213, A = 0 },
                    BandedFillColor = new ElementColor { R = 64, G = 113, B = 202, A = 0 },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle1Accent1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    FillColor = new ElementColor { R = 218, G = 227, B = 243, A = 0 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 68, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 68, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 68, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle2Accent1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 68, G = 114, B = 196, A = 0 },
                    FillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 91, G = 132, B = 203, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 91, G = 132, B = 203, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 91, G = 132, B = 203, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 91, G = 132, B = 203, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 91, G = 132, B = 203, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle3Accent1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    FillColor = new ElementColor { R = 218, G = 227, B = 243, A = 0 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 88, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 88, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 88, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = new ElementColor { R = 88, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 88, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 88, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 88, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle1Accent1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 68, G = 114, B = 196, A = 0 },
                    FillColor = new ElementColor { R = 233, G = 235, B = 245, A = 0 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 88, G = 129, B = 202, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 88, G = 129, B = 202, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 88, G = 129, B = 202, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 88, G = 129, B = 202, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 88, G = 129, B = 202, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle2Accent1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 68, G = 114, B = 196, A = 0 },
                    FillColor = new ElementColor { R = 207, G = 213, B = 234, A = 0 },
                    BandedFillColor = new ElementColor { R = 233, G = 235, B = 245, A = 0 },
                    TableBorder = new TableBorderStyle 
                    { 
                        Top = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }, 
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single } 
                    },
                    HeaderBorder = new TableBorderStyle 
                    { 
                        Top = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }, 
                        Right = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single } 
                    },
                    RowBorder = new TableBorderStyle 
                    { 
                        Top = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }, 
                        Right = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single } 
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle3Accent1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 68, G = 114, B = 196, A = 0 },
                    FillColor = new ElementColor { R = 231, G = 231, B = 231, A = 0 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 2, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle4Accent1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 233, G = 235, B = 245, A = 0 },
                    FillColor = new ElementColor { R = 207, G = 213, B = 234, A = 0 },
                    BandedFillColor = new ElementColor { R = 233, G = 235, B = 245, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 68, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 68, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 68, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = new ElementColor { R = 68, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 68, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 68, G = 114, B = 196, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.DarkStyle1Accent1, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 0, G = 0, B = 0, A = 0 },
                    FillColor = new ElementColor { R = 52, G = 89, B = 156, A = 0 },
                    BandedFillColor = new ElementColor { R = 68, G = 114, B = 196, A = 0 },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Accent 2 (Orange)
            {
                BuiltInTableStyle.ThemedStyle1Accent2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 237, G = 125, B = 49, A = 0 },
                    FillColor = new ElementColor { R = 243, G = 160, B = 114, A = 0 },
                    BandedFillColor = new ElementColor { R = 245, G = 178, B = 151, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 237, G = 128, B = 56, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 237, G = 128, B = 56, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = new ElementColor { R = 237, G = 128, B = 56, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 237, G = 128, B = 56, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.ThemedStyle2Accent2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 241, G = 139, B = 81, A = 0 },
                    FillColor = new ElementColor { R = 245, G = 157, B = 101, A = 0 },
                    BandedFillColor = new ElementColor { R = 245, G = 125, B = 45, A = 0 },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle1Accent2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    FillColor = new ElementColor { R = 251, G = 229, B = 214, A = 0 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 2, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle2Accent2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 237, G = 125, B = 49, A = 0 },
                    FillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 239, G = 141, B = 75, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 239, G = 141, B = 75, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 239, G = 141, B = 75, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 239, G = 141, B = 75, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 239, G = 141, B = 75, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = false,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.LightStyle3Accent2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    FillColor = new ElementColor { R = 251, G = 229, B = 214, A = 0 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle1Accent2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 237, G = 125, B = 49, A = 0 },
                    FillColor = new ElementColor { R = 252, G = 236, B = 232, A = 0 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle2Accent2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 237, G = 125, B = 49, A = 0 },
                    FillColor = new ElementColor { R = 248, G = 215, B = 205, A = 0 },
                    BandedFillColor = new ElementColor { R = 252, G = 236, B = 232, A = 0 },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle3Accent2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 237, G = 125, B = 49, A = 0 },
                    FillColor = new ElementColor { R = 231, G = 231, B = 231, A = 0 },
                    BandedFillColor = new ElementColor { R = 255, G = 255, B = 255, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 2, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 2, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 0, G = 0, B = 0, A = 0 }, Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.MediumStyle4Accent2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 252, G = 236, B = 232, A = 0 },
                    FillColor = new ElementColor { R = 248, G = 215, B = 205, A = 0 },
                    BandedFillColor = new ElementColor { R = 252, G = 236, B = 232, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Right = new TableLineStyle { Color = new ElementColor { R = 237, G = 125, B = 49, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },
            {
                BuiltInTableStyle.DarkStyle1Accent2, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 0, G = 0, B = 0, A = 0 },
                    FillColor = new ElementColor { R = 189, G = 98, B = 37, A = 0 },
                    BandedFillColor = new ElementColor { R = 237, G = 125, B = 49, A = 0 },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 2, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            },

            // Accent 3 (Gray)
            {
                BuiltInTableStyle.ThemedStyle1Accent3, new TableStyle
                {
                    HeadingFillColor = new ElementColor { R = 165, G = 165, B = 165, A = 0 },
                    FillColor = new ElementColor { R = 189, G = 189, B = 189, A = 0 },
                    BandedFillColor = new ElementColor { R = 201, G = 201, B = 201, A = 0 },
                    TableBorder = new TableBorderStyle
                    {
                        Top = new TableLineStyle { Color = new ElementColor { R = 169, G = 169, B = 169, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Left = new TableLineStyle { Color = new ElementColor { R = 169, G = 169, B = 169, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    HeaderBorder = new TableBorderStyle
                    {
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 255, G = 255, B = 255, A = 0 }, Width = 2, Style = LineStyle.Single }
                    },
                    RowBorder = new TableBorderStyle
                    {
                        Right = new TableLineStyle { Color = new ElementColor { R = 169, G = 169, B = 169, A = 0 }, Width = 1, Style = LineStyle.Single },
                        Bottom = new TableLineStyle { Color = new ElementColor { R = 169, G = 169, B = 169, A = 0 }, Width = 1, Style = LineStyle.Single }
                    },
                    BandedRows = true,
                    BandedColumns = false
                }
            }
        };
    }
}
