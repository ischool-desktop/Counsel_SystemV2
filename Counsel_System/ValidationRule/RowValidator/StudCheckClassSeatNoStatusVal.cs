using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;
using FISCA.Data;


namespace Counsel_System2.ValidationRule.RowValidator
{
    /// <summary>
    /// 檢查班級座號在同一狀態是否重複
    /// </summary>
    public class StudCheckClassSeatNoStatusVal:IRowVaildator
    {
        Dictionary<string, string> _StudStatusDict;

        public StudCheckClassSeatNoStatusVal()
        {
            _StudStatusDict = Utility.GetStudentStatusDBValDict();
        }

        public string Correct(IRowStream Value)
        {
            return string.Empty;
        }

        public string ToString(string template)
        {
            return template;
        }

        public bool Validate(IRowStream Value)
        {
            bool retVal = false;
            if (Value.Contains("班級") && Value.Contains("座號") && Value.Contains("狀態"))
            {
                string status = Value.GetValue("狀態").Trim();
                string key = Value.GetValue("班級") + "_" + Value.GetValue("座號") + "_";

                if (_StudStatusDict.ContainsKey(status))
                    key += _StudStatusDict[status];
                else
                    key += "1";

                if (Global._AllStudentClassSeatNoDictTemp.ContainsKey(key))
                    retVal = true;
            }

            return retVal;
        }
    }
}
