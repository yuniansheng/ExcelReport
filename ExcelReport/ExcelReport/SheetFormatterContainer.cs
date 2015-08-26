﻿/*
 类：SheetFormatterContainer
 描述：Sheet中元素的格式化器集合
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.Collections.Generic;

namespace ExcelReport
{
    public class SheetFormatterContainer
    {
        #region 成员字段及属性

        private string sheetName;

        public string SheetName
        {
            get { return sheetName; }
            protected set { sheetName = value; }
        }

        private IEnumerable<ElementFormatter> formatters;

        public IEnumerable<ElementFormatter> Formatters
        {
            get { return formatters; }
            protected set { formatters = value; }
        }

        #endregion 成员字段及属性

        /// 构造函数
        /// <param name="sheetName"></param>
        /// <param name="formatters"></param>
        public SheetFormatterContainer(string sheetName, IEnumerable<ElementFormatter> formatters)
        {
            this.sheetName = sheetName;
            this.formatters = formatters;
        }
    }
}