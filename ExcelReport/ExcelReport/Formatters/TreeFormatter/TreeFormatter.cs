/*
 类：TableFormatter
 描述：表格（元素）格式化器
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using ExcelReport.Formatters.TreeFormatter;
using System;
using System.Collections.Generic;

namespace ExcelReport
{
    public class TreeFormatter<TSource> : TableFormatter<TSource>
    {
        #region 成员字段

        private int _levelColumnIndex;
        //private Func<TSource, string> _getCode;
        private Func<TSource, int> _getLevel;
        private int _minLevel;

        #endregion 成员字段

        #region 属性

        #endregion

        /// 构造函数        
        public TreeFormatter(ParameterDictionary parameterDictionary, string parameterName, IEnumerable<TSource> dataSource)
            : base(parameterDictionary, parameterName, dataSource)
        {

        }

        /// 格式化操作
        public override void Format(SheetFormatterContext context)
        {
            context.ClearRowContent(TemplateRowIndex); //清除模板行单元格内容
            if (null == ColumnInfoList || ColumnInfoList.Count <= 0 || null == DataSource)
            {
                return;
            }
            var itemCount = 0;
            var sheet = context.Sheet;
            int rowIndex = 0;
            Stack<LevelInfo> levelStack = new Stack<LevelInfo>();


            foreach (TSource rowSource in DataSource)
            {
                if (itemCount++ > 0)
                {
                    //如果不是第一行时执行的操作
                    context.InsertEmptyRow(TemplateRowIndex);  //追加空行
                }

                //当前行索引
                rowIndex = context.GetCurrentRowIndex(TemplateRowIndex);
                //当前行的层级
                int level = GetLevel(rowSource);
                //当前行
                var row = sheet.GetRow(rowIndex);

                if (levelStack.Count > 0)
                {
                    LevelInfo top = levelStack.Peek();
                    while (level < top.Level)
                    {
                        sheet.GroupRow(top.RowIndex, rowIndex - 1);
                        levelStack.Pop();
                        top = levelStack.Peek();
                    }
                    if (level > top.Level)
                    {
                        levelStack.Push(new LevelInfo(rowIndex, level));
                    }
                }
                else
                {
                    levelStack.Push(new LevelInfo(rowIndex, level));
                }

                foreach (TableColumnInfo<TSource> colInfo in ColumnInfoList)
                {
                    var cell = row.GetCell(colInfo.ColumnIndex);
                    object value = colInfo.DgSetValue(rowSource);
                    if (colInfo.ColumnIndex == _levelColumnIndex)
                    {
                        //处理层级列，如果是层级列，则增加缩进
                        value = new string(' ', level * 2) + value;
                    }
                    SetCellValue(cell, value);
                }
            }

            while (levelStack.Count > 0)
            {
                LevelInfo top = levelStack.Pop();
                if (0 < top.Level)
                {
                    sheet.GroupRow(top.RowIndex, rowIndex);
                }
            }
        }

        /// <summary>
        /// 设置层级列的信息
        /// </summary>
        /// <param name="levelColumnIndex">作为层级列的列索引</param>
        /// <param name="getLevel">获取层级值的委托，引擎内部对Level=0的不增加缩进，为Level=1的记录增加一个缩进，依此类推</param>
        public void SetLevelColumn(int levelColumnIndex, Func<TSource, int> getLevel)
        {
            _levelColumnIndex = levelColumnIndex;
            _getLevel = getLevel;
            setMinLevel();
        }

        /// <summary>
        /// 设置层级列的信息
        /// </summary>
        /// <param name="levelColumnIndex">位于作为层级列的参数的参数名</param>
        /// <param name="getLevel">获取层级值的委托，引擎内部对Level=0的不增加缩进，为Level=1的记录增加一个缩进，依此类推</param>
        public void SetLevelColumn(string parameterName, Func<TSource, int> getLevel)
        {
            Parameter p = ParamDic[parameterName];
            if (p == null)
            {
                throw new ApplicationException("模板中未定义参数:" + parameterName + "，设置层级列出错！");
            }
            SetLevelColumn(p.CellPoint.Y, getLevel);
        }

        private void setMinLevel()
        {
            var enumerator = DataSource.GetEnumerator();
            if (enumerator.MoveNext())
            {
                _minLevel = _getLevel(enumerator.Current);
            }
        }

        private int GetLevel(TSource row)
        {
            //string levelCode = _getCode(row);
            //return levelCode.Length - levelCode.Replace(".", "").Length;
            return _getLevel(row) - _minLevel;
        }
    }
}