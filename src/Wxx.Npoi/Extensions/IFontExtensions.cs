using System;

namespace Wxx.Npoi.Extensions
{
    /// <summary>
    /// 
    /// </summary>
    public static class IFontExtensions
    {
        public static NPOI.SS.UserModel.IFont SetFontHeightInPoints(this NPOI.SS.UserModel.IFont font, short fontSize)
        {
            font.FontHeightInPoints = fontSize;
            return font;
        }

        public static NPOI.SS.UserModel.IFont SetColor(this NPOI.SS.UserModel.IFont font, short color)
        {
            font.Color = color;
            return font;
        }

        public static NPOI.SS.UserModel.IFont SetBoldweight(this NPOI.SS.UserModel.IFont font, short boldweight)
        {
            font.Boldweight = boldweight;
            return font;
        }
    }
}
