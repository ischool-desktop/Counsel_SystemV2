using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    //2016/9/8 穎驊仿作

    /// <summary>
    /// 身高及體重
    /// </summary>

   public  class CheckProcess10:ICheckProcess
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
            chkItems1.Add("身高");
            chkItems1.Add("體重");

            if (CheckDataTransfer.CheckSEMESTER_Error("本人概況", chkItems1, _Student)>0)
                _ErrorCount += 1;

            _TotalCount += 1;


            

            #endregion



        }

        public string GetMessage()
        {
            //2016/9/9 穎驊註解，經由與恩正討論，現在無論有缺漏，全部人的資料都要顯示出來，
            if (_ErrorCount > 0)
            {
                //return "未輸入完整：" + _ErrorCount + "/" + _TotalCount;
                return "" + (_TotalCount - _ErrorCount) + "/" + _TotalCount;
            }
            else
                //return "";
            return "" + (_TotalCount - _ErrorCount) + "/" + _TotalCount;
        }


        public void SetStudent(ClassStudent Student)
        {
            _Student = Student;
        }
    }
}
