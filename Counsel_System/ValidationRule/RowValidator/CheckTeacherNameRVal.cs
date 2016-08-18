using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;

namespace Counsel_System2.ValidationRule.RowValidator
{
    /// <summary>
    /// 檢查老師姓名是否存在
    /// </summary>
    public class CheckTeacherNameRVal:IRowVaildator
    {        
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
            if (Value.Contains("晤談老師"))
            { 
                // 表示有這位老師
                if(Global._AllTeacherNameIdDictTemp.ContainsKey(Value.GetValue("晤談老師")))
                    retVal =true ;            
            }
            return retVal;
        }
    }
}
