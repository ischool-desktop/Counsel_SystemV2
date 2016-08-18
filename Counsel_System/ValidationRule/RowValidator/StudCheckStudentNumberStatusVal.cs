using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;
using FISCA.Data;

namespace Counsel_System2.ValidationRule.RowValidator
{
    /// <summary>
    /// 檢查學號在不同狀態下是否重複
    /// </summary>
    public class StudCheckStudentNumberStatusVal : IRowVaildator
    {
        Dictionary<string, string> _StudStatusDict;
        public StudCheckStudentNumberStatusVal()
        {
            _StudStatusDict = Utility.GetStudentStatusDBValDict();
        }

        #region IRowVaildator 成員

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
            if (Value.Contains("學號") && Value.Contains("狀態"))
            {
                string status = Value.GetValue("狀態").Trim();
                string key = Value.GetValue("學號") + "_";

                if (_StudStatusDict.ContainsKey(status))
                    key += _StudStatusDict[status];
                else
                    key += "1";

                if (Global._AllStudentNumberStatusIDTemp.ContainsKey(key))
                    retVal = true;
            }

            return retVal;
        }

        #endregion
    }
}
