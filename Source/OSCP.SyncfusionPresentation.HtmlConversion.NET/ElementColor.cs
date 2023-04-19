using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal struct ElementColor
    {
        internal byte R;
        internal byte G;
        internal byte B;
        internal byte A;
        internal string Css
        {
            get
            {
                return $"rgba({this.R},{this.G},{this.B},{(this.A == 0 ? 1 : (100F - this.A) / 100F)})";
            }
        }
    }
}
