using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;


/*
 * 2015-03-12 18:47 问题：
 *     命名空间不知道为什么不能用Wxx.Infrastructure.NPOI.Excel，即Npoi换成NPOI，当时环境：.net4；如果用了
 *     在使用NPOI中的类时如果使用完整类名（NPOI.SS.UserModel.IWorkbook）就会编译不通过。
 */
namespace Wxx.Npoi
{
    /// <summary>
    /// NPOI对Excel操作的帮助类
    /// </summary>
    public static class NpoiExcelHelper
    {
        /// <summary>
        /// 创建指定格式的工作簿
        /// </summary>
        /// <param name="excelFormat">格式类型</param>
        /// <returns></returns>
        public static NPOI.SS.UserModel.IWorkbook GetWorkbook(ExcelFormat excelFormat)
        {
            NPOI.SS.UserModel.IWorkbook workbook = null;
            switch (excelFormat)
            {
                case ExcelFormat.Xls:
                    // 2003
                    workbook = new NPOI.HSSF.UserModel.HSSFWorkbook();
                    break;
                case ExcelFormat.Xlsx:
                    // 2007
                    workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
                    break;
                default:
                    break;
            }
            return workbook;
        }

        /// <summary>
        /// 打开工作簿
        /// </summary>
        /// <param name="inputStream">.xls或.xlsx文件流</param>
        /// <returns></returns>
        public static NPOI.SS.UserModel.IWorkbook GetWorkbook(Stream inputStream)
        {
            // 注意：SS.UserModel.WorkbookFactory 需要添加“NPOI.OpenXml4Net.dll”的引用

            // WorkbookFactory内部可以通过inputStream判断是2003、还是2007。
            return NPOI.SS.UserModel.WorkbookFactory.Create(inputStream);
        }

        /// <summary>
        /// 打开工作簿
        /// </summary>
        /// <param name="file">.xls或.xlsx文件的物理路径</param>
        /// <returns></returns>
        public static NPOI.SS.UserModel.IWorkbook GetWorkbook(string file)
        {
            return NPOI.SS.UserModel.WorkbookFactory.Create(file);
        }
    }
}
