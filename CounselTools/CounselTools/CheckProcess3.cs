using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    /// <summary>
    /// 學習狀況
    /// </summary>
   public  class CheckProcess3:ICheckProcess
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
            #region SEMESTER
            List<string> chkItems1 = new List<string>();
            chkItems1.Add("社團幹部");
            chkItems1.Add("班級幹部");
            if (CheckDataTransfer.CheckSEMESTER_Error(_GroupName, chkItems1, _Student)>0)
                _ErrorCount += 1;

            _TotalCount += 1;

            #endregion

            #region YEARLY
            List<string> chkItems2 = new List<string>();
            chkItems2.Add("休閒興趣");
            chkItems2.Add("特殊專長");
            chkItems2.Add("最喜歡的學科");
            chkItems2.Add("最感困難的學科");
            if (CheckDataTransfer.CheckYEARLY_Error(_GroupName, chkItems2, _Student)>0)
                _ErrorCount+=1;
            _TotalCount+=1;

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
