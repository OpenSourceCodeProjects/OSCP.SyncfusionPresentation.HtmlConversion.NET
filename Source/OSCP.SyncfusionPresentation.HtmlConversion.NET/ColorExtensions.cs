using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Syncfusion.Presentation;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal static class ColorExtensions
    {
        internal static string ToCss(this Syncfusion.Drawing.Color color)
        {
            float alpha = (color.A == 0 || color.A == 255)
                ? 1F
                : (100F - ((color.A / 255F) * 100F)) / 100F;

            return $"rgba({color.R},{color.G},{color.B},{alpha})";
        }

        internal static ColorObject FromCss(string styleValue)
        {
            ColorObject colorObject = null;

            // Try rgba
            Match rgbaMatch = Regex.Match(styleValue, @"rgba\((\d{1,3}),\s?(\d{1,3}),\s?(\d{1,3}),\s?(\d?.?\d{1,2}?)\)");
            if (rgbaMatch.Success == true && rgbaMatch.Groups.Count == 5)
            {
                float a = float.Parse(rgbaMatch.Groups[4].Value);
                int alpha = (int)(a == 1F 
                    ? 255
                    : (1F - float.Parse(rgbaMatch.Groups[4].Value)) * 255);

                colorObject = ColorObject.FromArgb(
                    alpha,
                    int.Parse(rgbaMatch.Groups[1].Value),
                    int.Parse(rgbaMatch.Groups[2].Value),
                    int.Parse(rgbaMatch.Groups[3].Value)
                ) as ColorObject;
            }
            else
            {
                // Try rgb.
                Match rgbMatch = Regex.Match(styleValue, @"rgb\((\d{1,3}),\s?(\d{1,3}),\s?(\d{1,3})\)");
                if (rgbMatch.Success == true && rgbMatch.Groups.Count == 4)
                {
                    colorObject = ColorObject.FromArgb(
                        int.Parse(rgbMatch.Groups[1].Value),
                        int.Parse(rgbMatch.Groups[2].Value),
                        int.Parse(rgbMatch.Groups[3].Value)
                    ) as ColorObject;
                }
            }

            return colorObject;
        }
    }
}
