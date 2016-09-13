using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    /// <summary>
    /// 畢業後規劃
    /// </summary>
    public class CheckProcess7:ICheckProcess
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
            #region MULTI_ANSWER

            List<string> chkItems1 = new List<string>();
            chkItems1.Add("升學意願");
            chkItems1.Add("受訓地區");
            chkItems1.Add("參加職業訓練");
            chkItems1.Add("就業意願");
            _ErrorCount += CheckDataTransfer.CheckMULTI_ANSWER_Error("畢業後計畫", chkItems1, _Student);
            _TotalCount += chkItems1.Count;

            #endregion

            #region PRIORITY

            List<string> chkItems2 = new List<string>();
            chkItems2.Add("將來職業");
            chkItems2.Add("就業地區");
            // 這是一項
            if (CheckDataTransfer.CheckPRIORITY_Error("畢業後計畫", chkItems2, _Student) > 0)
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
