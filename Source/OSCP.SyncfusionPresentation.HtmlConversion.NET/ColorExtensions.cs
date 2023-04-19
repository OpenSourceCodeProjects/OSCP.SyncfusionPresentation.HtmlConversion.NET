using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal static class ColorExtensions
    {
        internal static string ToCss(this Syncfusion.Drawing.Color color)
        {
            float alpha = color.A == 0
                ? 1F
                : (
                    color.A == 1
                        ? 0
                        : (100F - ((color.A / 255F) * 100F)) / 100F
                );

            if (color.A != 0 && color.A != 1) 
            {
                int i = 1;
            }

            return $"rgba({color.R},{color.G},{color.B},{alpha})";
        }
    }
}
