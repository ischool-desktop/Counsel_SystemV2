using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
 
    /// <summary>
    /// 學習
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
            #region YEARLY
            List<string> chkItems2 = new List<string>();
            chkItems2.Add("休閒興趣");
            chkItems2.Add("特殊專長");
            chkItems2.Add("最喜歡的學科");
            chkItems2.Add("最感困難的學科");
            chkItems2.Add("特殊專長_樂器演奏");
            chkItems2.Add("特殊專長_外語能力");


            _ErrorCount += CheckDataTransfer.CheckYEARLY_Error("學習狀況", chkItems2, _Student);

            _TotalCount = chkItems2.Count;

            //if (CheckDataTransfer.CheckYEARLY_Error(_GroupName, chkItems2, _Student) > 0)
            //    _ErrorCount += 1;
            //_TotalCount += 1;

            #endregion


        }

        public string GetMessage()
        {
            //2016/9/9 穎驊註解，經由與恩正討論，現在無論有缺漏，全部人的資料都要顯示出來，
            if (_ErrorCount > 0)
            {
                //return "未輸入完整：" + _ErrorCount + "/" + _TotalCount;
                return "輸入況狀：" + (_TotalCount - _ErrorCount) + "/" + _TotalCount;
            }
            else
                //return "";
                return "輸入況狀：" + (_TotalCount - _ErrorCount) + "/" + _TotalCount;
        }


        public void SetStudent(ClassStudent Student)
        {
            _Student = Student;
        }
    }
}
