using System;

namespace Wxx.Npoi.Extensions
{
    /// <summary>
    /// 
    /// </summary>
    public static class IRowExtensions
    {
        /// <summary>
        /// 获取或创建单元格
        /// </summary>
        public static NPOI.SS.UserModel.ICell GetOrCreateCell(this NPOI.SS.UserModel.IRow row, int cellnum)
        {
            return row.GetCell(cellnum) ?? row.CreateCell(cellnum);
        }

        public static NPOI.SS.UserModel.IRow GetOrCreateCell(this NPOI.SS.UserModel.IRow row, int cellnum, Action<NPOI.SS.UserModel.ICell> action)
        {
            var cell = row.GetOrCreateCell(cellnum);
            if (action != null)
                action(cell);
            return row;
        }
    }
}
