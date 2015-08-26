﻿using NPOI.SS.UserModel;
/*
 类：ExportHelper
 描述：导出助手类
 
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/
using System;
using System.IO;
using System.Web;

namespace ExcelReport
{
    public static partial class ExportHelper
    {

        /// 导出到本地
        /// <param name="templateFile"></param>
        /// <param name="targetFile"></param>
        /// <param name="containers"></param>
        public static void ExportToLocal(string templateFile, string targetFile, params SheetFormatterContainer[] containers)
        {
            #region 参数验证

            if (string.IsNullOrEmpty(templateFile))
            {
                throw new ArgumentNullException("templateFile");
            }
            if (!File.Exists(templateFile))
            {
                throw new FileNotFoundException(templateFile + " 文件不存在!");
            }

            #endregion 参数验证

            IWorkbook workbook = NPOIHelper.LoadWorkbook(templateFile);
            ExportToLocal(workbook, targetFile, containers);
        }

        public static void ExportToLocal(IWorkbook workbook, string targetFile, params SheetFormatterContainer[] containers)
        {
            #region 参数验证

            if (string.IsNullOrEmpty(targetFile))
            {
                throw new ArgumentNullException("targetFile");
            }

            #endregion 参数验证

            using (FileStream fs = File.OpenWrite(targetFile))
            {
                var buffer = Export.ExportToBuffer(workbook, containers);
                fs.Write(buffer, 0, buffer.Length);
                fs.Flush();
            }
        }

        /// 导出到Web
        /// <param name="templateFile"></param>
        /// <param name="targetFile"></param>
        /// <param name="containers"></param>
        public static void ExportToWeb(string templateFile, string targetFile, params SheetFormatterContainer[] containers)
        {
            #region 参数验证

            if (string.IsNullOrEmpty(templateFile))
            {
                throw new ArgumentNullException("templateFile");
            }
            if (!File.Exists(templateFile))
            {
                throw new FileNotFoundException(templateFile + " 文件不存在!");
            }

            #endregion 参数验证

            IWorkbook workbook = NPOIHelper.LoadWorkbook(templateFile);
            ExportToWeb(workbook, targetFile, containers);
        }

        public static void ExportToWeb(IWorkbook workbook, string targetFile, params SheetFormatterContainer[] containers)
        {
            #region 参数验证

            if (string.IsNullOrEmpty(targetFile))
            {
                throw new ArgumentNullException("targetFile");
            }

            #endregion 参数验证

            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
            HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(targetFile, System.Text.Encoding.UTF8));
            HttpContext.Current.Response.BinaryWrite(Export.ExportToBuffer(workbook, containers));
            HttpContext.Current.Response.End();
        }
    }
}
