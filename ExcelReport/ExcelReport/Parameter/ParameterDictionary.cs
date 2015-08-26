using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReport
{
    public class ParameterDictionary
    {
        private Dictionary<string, Parameter> parameters;

        public Parameter this[string parameterName]
        {
            get
            {
                Parameter p;
                if (parameters.TryGetValue(parameterName, out p))
                {
                    return p;
                }
                return null;
            }
            set
            {
                parameters[parameterName] = value;
            }
        }

        public ParameterDictionary()
        {
            parameters = new Dictionary<string, Parameter>();
        }
    }
}
