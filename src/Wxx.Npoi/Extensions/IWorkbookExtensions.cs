using System;
using System.Collections.Generic;

namespace Wxx.Npoi.Extensions
{
    /// <summary>
    /// 
    /// </summary>
    public static class IWorkbookExtensions
    {
        /// <summary>
        /// 获取Excel格式；例如：2003、2007等
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static ExcelFormat GetExcelFormat(this NPOI.SS.UserModel.IWorkbook workbook)
        { 
            ExcelFormat format = ExcelFormat.None;
            if ((workbook as NPOI.HSSF.UserModel.HSSFWorkbook) == null)
                format = ExcelFormat.Xls;
            else if ((workbook as NPOI.XSSF.UserModel.XSSFWorkbook) == null)
                format = ExcelFormat.Xlsx;
            return format;
        }

        #region 创建或获取工作页（sheet）
        /// <summary>
        /// 创建或获取工作页
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name">sheet名称</param>
        /// <returns></returns>
        public static NPOI.SS.UserModel.ISheet GetOrCreateSheet(this NPOI.SS.UserModel.IWorkbook workbook, string name)
        {
            return workbook.GetSheet(name) ?? workbook.CreateSheet(name);
        }

        public static NPOI.SS.UserModel.IWorkbook GetOrCreateSheet(this NPOI.SS.UserModel.IWorkbook workbook, string name, Action<NPOI.SS.UserModel.ISheet> action)
        {
            var sheet = workbook.GetOrCreateSheet(name);
            if (action != null)
                action(sheet);
            return workbook;
        }

        /// <summary>
        /// 创建或获取工作页
        /// <para>
        /// 获取第一页，如果不存在则创建
        /// </para>
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static NPOI.SS.UserModel.ISheet GetOrCreateSheet(this NPOI.SS.UserModel.IWorkbook workbook)
        {
            NPOI.SS.UserModel.ISheet sheet = null;
            if (workbook.NumberOfSheets > 0)
                sheet = workbook.GetSheetAt(0);
            return sheet ?? workbook.CreateSheet();
        }

        public static NPOI.SS.UserModel.IWorkbook GetOrCreateSheet(this NPOI.SS.UserModel.IWorkbook workbook, Action<NPOI.SS.UserModel.ISheet> action)
        {
            var sheet = workbook.GetOrCreateSheet();
            if (action != null)
                action(sheet);
            return workbook;
        }

        /// <summary>
        /// 获取workbook所有的sheet
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static List<NPOI.SS.UserModel.ISheet> GetSheets(this NPOI.SS.UserModel.IWorkbook workbook)
        {
            List<NPOI.SS.UserModel.ISheet> res = new List<NPOI.SS.UserModel.ISheet>();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(i);
                // sheet不为空且未隐藏
                if (sheet != null && !workbook.IsSheetHidden(i))
                    res.Add(sheet);
            }
            return res;
        }
        #endregion


        #region 获取或创建数据格式
        /// <summary>
        /// 获取或创建数据格式
        /// <para>
        /// 参考：http://tonyqus.sinaapp.com/?p=108
        /// </para>
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="format">格式字符串，如：yyyy-mm-dd</param>
        /// <returns>返回格式所对应的索引</returns>
        public static short GetOrCreateDataFormat(this NPOI.SS.UserModel.IWorkbook workbook, string format)
        {
            return workbook.CreateDataFormat().GetFormat(format);
        }

        /// <summary>
        /// 创建日期样式
        /// </summary>
        public static short CreateDateFormat(this NPOI.SS.UserModel.IWorkbook workbook, string format = "yyyy-mm-dd")
        {
            return workbook.GetOrCreateDataFormat(format);
        }
        #endregion
        
    }
}
