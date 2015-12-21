using System;

namespace Wxx.Npoi.Extensions
{
    /// <summary>
    /// 
    /// </summary>
    public static class ISheetExtensions
    {
        #region 获取或创建行
        /// <summary>
        /// 获取或创建行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        public static NPOI.SS.UserModel.IRow GetOrCreateRow(this NPOI.SS.UserModel.ISheet sheet, int rownum)
        {
            return sheet.GetRow(rownum) ?? sheet.CreateRow(rownum);
        }

        public static NPOI.SS.UserModel.ISheet GetOrCreateRow(this NPOI.SS.UserModel.ISheet sheet, int rownum, Action<NPOI.SS.UserModel.IRow> action)
        {
            var row = sheet.GetOrCreateRow(rownum);
            if(action != null)
                action(row);
            return sheet;
        }
        #endregion

        #region 合并单元格
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="startRowIndex">起始行索引</param>
        /// <param name="endRowIndex">结束行索引</param>
        /// <param name="startColumnIndex">起始列索引</param>
        /// <param name="endColumnIndex">结束列索引</param>
        /// <returns>合并区域的索引</returns>
        public static NPOI.SS.Util.CellRangeAddress MergeCell(this NPOI.SS.UserModel.ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
        {
            var region = new NPOI.SS.Util.CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            sheet.AddMergedRegion(region);
            return region;
        }


        public static void SetMergeCellBorder(this NPOI.SS.UserModel.ISheet sheet, NPOI.SS.UserModel.BorderStyle borderStyle, int color, int firstRow, int lastRow, int firstCol, int lastCol)
        {
            /*
            var region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            NPOI.HSSF.Util.HSSFRegionUtil.SetBorderBottom(borderStyle, region, sheet, 
            int index = -1;
            var region = new NPOI.SS.Util.CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            index = sheet.AddMergedRegion(region);
            return index;*/

        }
        #endregion
    }
}
