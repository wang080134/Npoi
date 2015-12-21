using System;

namespace Wxx.Npoi.Extensions
{
    /// <summary>
    /// ICell扩展
    /// </summary>
    public static class ICellExtensions
    {
        #region 设置单元格的值
        /// <summary>
        /// 设置单元格的值
        /// </summary>
        public static void SetCellValueExt(this NPOI.SS.UserModel.ICell cell, object value)
        {
            if (value == null)
                return;
            switch (value.GetType().ToString())
            {
                case "System.String":
                    cell.SetCellValue(value.ToString());
                    break;
                case "System.DateTime":
                    cell.SetCellValue(CastHelper.CastTo<DateTime>(value));
                    cell.SetCellStyle(style => style.SetDataFormat(cell.Row.Sheet.Workbook.CreateDateFormat()));
                    break;
                case "System.Boolean":
                    cell.SetCellValue(CastHelper.CastTo<bool>(value));
                    break;
                case "System.Byte":
                case "System.Int16":
                case "System.Int32":
                case "System.Int64":
                    cell.SetCellValue(CastHelper.CastTo<int>(value));
                    break;
                case "System.Double":
                case "System.Decimal":
                    cell.SetCellValue(CastHelper.CastTo<double>(value));
                    break;
                case "System.DBNull":
                default:
                    cell.SetCellValue("");
                    break;
            }
        }
        #endregion

        #region 获取单元格的值
        /// <summary>
        /// 设置单元格的值
        /// </summary>
        public static object GetCellValue(this NPOI.SS.UserModel.ICell cell)
        {
            object value = null;
            switch (cell.CellType)
            {
                // 空单元格
                case NPOI.SS.UserModel.CellType.Blank:
                    // value = DBNull.Value;
                    break;
                // 
                case NPOI.SS.UserModel.CellType.Boolean:
                    value = cell.BooleanCellValue;
                    break;
                // 
                case NPOI.SS.UserModel.CellType.Error:
                    value = cell.ErrorCellValue;
                    break;
                // 
                case NPOI.SS.UserModel.CellType.Formula:
                    value = cell.StringCellValue;
                    break;
                // 数值单元格
                case NPOI.SS.UserModel.CellType.Numeric:
                    value = cell.NumericCellValue;
                    break;
                // 
                case NPOI.SS.UserModel.CellType.String:
                    value = cell.StringCellValue;
                    break;
                // 
                case NPOI.SS.UserModel.CellType.Unknown:
                    //value = cell.CellType
                    break;
                default:
                    break;
            }
            return value;
        }

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        public static T GetCellValue<T>(this NPOI.SS.UserModel.ICell cell)
        {
            return CastHelper.CastTo<T>(cell.GetCellValue());
        }

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        public static T GetCellValue<T>(this NPOI.SS.UserModel.ICell cell, T deafultValue)
        {
            return CastHelper.CastTo<T>(cell.GetCellValue(), deafultValue);
        }
        #endregion

        public static NPOI.SS.UserModel.ICell SetCellStyle(this NPOI.SS.UserModel.ICell cell, NPOI.SS.UserModel.ICellStyle cellStyle)
        {
            cell.CellStyle = cellStyle;
            return cell;
        }

        public static NPOI.SS.UserModel.ICell SetCellStyle(this NPOI.SS.UserModel.ICell cell, Action<NPOI.SS.UserModel.ICellStyle> action)
        {
            var cellStyle = cell.Row.Sheet.Workbook.CreateCellStyle();
            if (cell.CellStyle != null)
                cell.CellStyle.CloneStyleFrom(cellStyle);
            if (action != null)
                action(cellStyle);
            cell.CellStyle = cellStyle;
            return cell;
        }
    }
}
