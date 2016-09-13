using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    //2016/9/8 穎驊仿作

    /// <summary>
    /// 監護人資料
    /// </summary>
    
    public class CheckProcess11:ICheckProcess
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
            chkItems1.Add("監護人_姓名");
            chkItems1.Add("監護人_性別");
            chkItems1.Add("監護人_關係");
            chkItems1.Add("監護人_電話");
            chkItems1.Add("監護人_通訊地址");
         

            // SINGLE_ANSWER
            _ErrorCount += CheckDataTransfer.CheckSINGLE_ANSWER_Error("家庭狀況", chkItems1, _Student);
            _TotalCount += chkItems1.Count;
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
