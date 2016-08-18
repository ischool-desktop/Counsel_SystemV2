using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    /// <summary>
    /// 備註
    /// </summary>
    public class CheckProcess8:ICheckProcess
    {
        string _GroupName;
        ClassStudent _Student;
        int _ErrorCount = 0, _TotalCount = 0;
        Dictionary<string, string> _ErrorDict = new Dictionary<string, string>();

        public void SetGroupName(string GroupName)
        {
            _GroupName = GroupName;
        }

     
        public Dictionary<string, string> GetErrorData()
        {
            return _ErrorDict;
        }

        public int GetErrorCount()
        {
            return _ErrorCount;
        }

        public int GetTotalCount()
        {
            return _TotalCount;
        }

        public void Start()
        {
            #region SINGLE_ANSWER
            List<string> chkItems1 = new List<string>();
            chkItems1.Add("備註");

            // SINGLE_ANSWER
            _ErrorCount += CheckDataTransfer.CheckSINGLE_ANSWER_Error(_GroupName, chkItems1, _Student);
            _TotalCount += chkItems1.Count;
            #endregion
        }

        public string GetMessage()
        {
            if (_ErrorCount > 0)
            {
                return "未輸入完整：" + _ErrorCount + "/" + _TotalCount;
            }
            else
                return "";
        }


        public void SetStudent(ClassStudent Student)
        {
            _Student = Student;
        }
    }
}
