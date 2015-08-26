using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReport
{
    public class Sheet2ParameterDictionary
    {
        private Dictionary<string, ParameterDictionary> _container;
        private IWorkbook _workbook;

        public ParameterDictionary this[string sheetName]
        {
            get
            {
                ParameterDictionary paramDic;
                _container.TryGetValue(sheetName, out paramDic);
                return paramDic;
            }
        }

        public ParameterDictionary this[int sheetIndex]
        {
            get
            {
                string sheetName = _workbook.GetSheetName(sheetIndex);
                return this[sheetName];
            }
        }

        internal Sheet2ParameterDictionary(IWorkbook workbook)
        {
            _container = new Dictionary<string, ParameterDictionary>();
            _workbook = workbook;
        }

        public void AddParameter(Parameter parameter)
        {
            ParameterDictionary paramDic;
            if (!_container.TryGetValue(parameter.SheetName, out paramDic))
            {
                paramDic = new ParameterDictionary();
                _container[parameter.SheetName] = paramDic;
            }
            paramDic[parameter.ParameterName] = parameter;
        }
    }
}
