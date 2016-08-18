using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    /// <summary>
    /// 適應情形
    /// </summary>
    public class CheckProcess9:ICheckProcess
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
            #region YEARLY
            List<string> chkItems1 = new List<string>();
            chkItems1.Add("人生態度");
            chkItems1.Add("人際關係");
            chkItems1.Add("內向行為");
            chkItems1.Add("外向行為");
            chkItems1.Add("生活習慣");
            chkItems1.Add("服務熱忱");
            chkItems1.Add("學習動機");

            // 這屬於一項
            if (CheckDataTransfer.CheckYEARLY_Error(_GroupName, chkItems1, _Student) > 0)
                _ErrorCount += 1;
            
            _TotalCount += 1;

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
