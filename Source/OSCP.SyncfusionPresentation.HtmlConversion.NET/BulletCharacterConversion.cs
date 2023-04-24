using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OSCP.SyncfusionPresentation.HtmlConversion.NET
{
    internal class BulletCharacterConversion
    {
        // http://www.alanwood.net/demos/wingdings.html
        private static readonly Dictionary<int, string> WINDINGS_DEC_TO_UNICODE_HEX = new Dictionary<int, string>()
        {
            { 113, "\"\\2751\"" },  // boxshadowdwn
            { 118, "\"\\2756\"" },  // xrhombus
            { 167, "square" },      // square4
            { 216, "\"\\2B9A\"" },  // head2right
            { 252, "\"\\2714\"" }   // checkbld
        };

        public static string BulletCharacterToListStyleType(string fontName, char bulletCharacter)
        {
            string listStyleType = string.Empty;

            switch (bulletCharacter)
            {
                case '•':
                    listStyleType = "disc";
                    break;
                case 'o':
                    listStyleType = "circle";
                    break;
                default:
                    if (fontName == "Wingdings")
                    {
                        WINDINGS_DEC_TO_UNICODE_HEX.TryGetValue(Convert.ToInt32(bulletCharacter), out listStyleType);
                    }
                    if (listStyleType == null || listStyleType.Length == 0)
                    {
                        listStyleType = $"\"{bulletCharacter}\"";
                    }

                    break;
            }

            return listStyleType;
        }

        public static char ListStyleTypeToBulletCharacter(string fontName, string listStyleType)
        {
            char bulletCharacter = ' ';

            switch (listStyleType)
            {
                case "disc":
                    bulletCharacter = '•';
                    break;
                case "circle":
                    bulletCharacter = 'o';
                    break;
                default:
                    if (fontName == "Wingdings")
                    {
                        foreach (int key in WINDINGS_DEC_TO_UNICODE_HEX.Keys)
                        {
                            if (listStyleType == WINDINGS_DEC_TO_UNICODE_HEX[key])
                            {
                                bulletCharacter = Convert.ToChar(key);
                                break;
                            }
                        }
                    }
                    if (bulletCharacter == ' ' && char.TryParse(listStyleType, out bulletCharacter) == false)
                    {
                        bulletCharacter = '•';
                    }

                    break;
            }

            return bulletCharacter;
        }
    }
}
