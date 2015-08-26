using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReport.Formatters.TreeFormatter
{
    internal class LevelInfo
    {
        public int RowIndex { get; set; }
        public int Level { get; set; }

        public LevelInfo()
        {

        }

        public LevelInfo(int rowIndex,int level)
        {
            RowIndex = rowIndex;
            Level = level;
        }
    }
}
