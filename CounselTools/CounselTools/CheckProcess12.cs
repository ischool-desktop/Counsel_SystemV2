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
    /// 
    /// <summary>
    /// 幹部資訊
    /// </summary>
   public  class CheckProcess12:ICheckProcess
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
            if (CheckDataTransfer.CheckSEMESTER_Error("學習狀況", chkItems1, _Student)>0)
                _ErrorCount += 1;

            _TotalCount += 1;

            #endregion


            // 2016/9/8 穎驊註解，因應文華專屬的輔導系統，在既有的項目將會做調整
            //原名:學期狀況表格 底下包括下面的項目，新版不需要了 故註解掉

            //#region YEARLY
            //List<string> chkItems2 = new List<string>();
            //chkItems2.Add("休閒興趣");
            //chkItems2.Add("特殊專長");
            //chkItems2.Add("最喜歡的學科");
            //chkItems2.Add("最感困難的學科");
            //if (CheckDataTransfer.CheckYEARLY_Error(_GroupName, chkItems2, _Student)>0)
            //    _ErrorCount+=1;
            //_TotalCount+=1;

            //#endregion


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
