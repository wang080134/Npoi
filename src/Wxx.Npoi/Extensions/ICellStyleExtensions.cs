using System;

namespace Wxx.Npoi.Extensions
{
    /// <summary>
    /// 
    /// </summary>
    public static class ICellStyleExtensions
    {
        // 水平对齐
        public static NPOI.SS.UserModel.ICellStyle SetDataFormat(this NPOI.SS.UserModel.ICellStyle cellStyle, short dataFormat)
        {
            cellStyle.DataFormat = dataFormat;
            return cellStyle;
        }

        // 水平对齐
        public static NPOI.SS.UserModel.ICellStyle SetHorizontalAlignment(this NPOI.SS.UserModel.ICellStyle cellStyle, NPOI.SS.UserModel.HorizontalAlignment horizontalAlignment)
        {
            cellStyle.Alignment = horizontalAlignment;
            return cellStyle;
        }

        // 垂直对齐
        public static NPOI.SS.UserModel.ICellStyle SetVerticalAlignment(this NPOI.SS.UserModel.ICellStyle cellStyle, NPOI.SS.UserModel.VerticalAlignment verticalAlignment)
        {
            cellStyle.VerticalAlignment = verticalAlignment;
            return cellStyle;
        }

        // 
        public static NPOI.SS.UserModel.ICellStyle SetFillForegroundColor(this NPOI.SS.UserModel.ICellStyle cellStyle, short fillForegroundColor)
        {
            cellStyle.FillForegroundColor = fillForegroundColor;
            return cellStyle;
        }

        // 填充模式
        public static NPOI.SS.UserModel.ICellStyle SetFillPattern(this NPOI.SS.UserModel.ICellStyle cellStyle, NPOI.SS.UserModel.FillPattern fillPattern)
        {
            cellStyle.FillPattern = fillPattern;
            return cellStyle;
        }

        #region 边框样式
        
        public static NPOI.SS.UserModel.ICellStyle SetBorderTop(this NPOI.SS.UserModel.ICellStyle cellStyle, NPOI.SS.UserModel.BorderStyle borderStyle)
        {
            cellStyle.BorderTop = borderStyle;
            return cellStyle;
        }
        public static NPOI.SS.UserModel.ICellStyle SetBorderRight(this NPOI.SS.UserModel.ICellStyle cellStyle, NPOI.SS.UserModel.BorderStyle borderStyle)
        {
            cellStyle.BorderRight = borderStyle;
            return cellStyle;
        }
        public static NPOI.SS.UserModel.ICellStyle SetBorderBottom(this NPOI.SS.UserModel.ICellStyle cellStyle, NPOI.SS.UserModel.BorderStyle borderStyle)
        {
            cellStyle.BorderBottom = borderStyle;
            return cellStyle;
        }
        public static NPOI.SS.UserModel.ICellStyle SetBorderLeft(this NPOI.SS.UserModel.ICellStyle cellStyle, NPOI.SS.UserModel.BorderStyle borderStyle)
        {
            cellStyle.BorderLeft = borderStyle;
            return cellStyle;
        }

        public static NPOI.SS.UserModel.ICellStyle SetTopBorderColor(this NPOI.SS.UserModel.ICellStyle cellStyle, short borderColor)
        {
            cellStyle.TopBorderColor = borderColor;
            return cellStyle;
        }
        public static NPOI.SS.UserModel.ICellStyle SetRightBorderColor(this NPOI.SS.UserModel.ICellStyle cellStyle, short borderColor)
        {
            cellStyle.RightBorderColor = borderColor;
            return cellStyle;
        }
        public static NPOI.SS.UserModel.ICellStyle SetBottomBorderColor(this NPOI.SS.UserModel.ICellStyle cellStyle, short borderColor)
        {
            cellStyle.BottomBorderColor = borderColor;
            return cellStyle;
        }
        public static NPOI.SS.UserModel.ICellStyle SetLeftBorderColor(this NPOI.SS.UserModel.ICellStyle cellStyle, short borderColor)
        {
            cellStyle.LeftBorderColor = borderColor;
            return cellStyle;
        }

        #endregion

        public static NPOI.SS.UserModel.ICellStyle SetFont(this NPOI.SS.UserModel.ICellStyle cellStyle, NPOI.SS.UserModel.IWorkbook workBook, Action<NPOI.SS.UserModel.IFont> action)
        {
            NPOI.SS.UserModel.IFont font = workBook.CreateFont();
            if (action != null)
                action(font);
            cellStyle.SetFont(font);
            return cellStyle;
        }
    }
}
