using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    /// <summary>
    /// 家庭訊息
    /// </summary>
    public class CheckProcess14:ICheckProcess
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
            List<string> chkItems4 = new List<string>();
            chkItems4.Add("父母關係");
            chkItems4.Add("父親管教方式");
            chkItems4.Add("本人住宿");
            chkItems4.Add("母親管教方式");
            chkItems4.Add("我覺得是否足夠");
            chkItems4.Add("每星期零用錢");
            chkItems4.Add("居住環境");
            chkItems4.Add("家庭氣氛");
            chkItems4.Add("經濟狀況");


            _ErrorCount += CheckDataTransfer.CheckYEARLY_Error("家庭狀況", chkItems4, _Student);

            _TotalCount += chkItems4.Count;

            //// 這算一項
            //if (CheckDataTransfer.CheckYEARLY_Error("家庭狀況", chkItems4, _Student) > 0)
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
