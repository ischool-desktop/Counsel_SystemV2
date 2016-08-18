using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;

namespace Counsel_System2.ValidationRule.RowValidator
{
    public class StudDataCheckQuizDataStudentNumberRVal : IRowVaildator
    {
        public StudDataCheckQuizDataStudentNumberRVal()
        {
            List<string> id = Global._HasStudQuizDataDictTemp.Keys.ToList();
        
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
            bool retVal = true;
            if (Value.Contains("學號"))
            {
                if(Global._HasStudQuizDataDictTemp.ContainsKey(Value.GetValue("學號")))
                    retVal= false ;
            }

            return retVal;
        }

        #endregion
    }
}
