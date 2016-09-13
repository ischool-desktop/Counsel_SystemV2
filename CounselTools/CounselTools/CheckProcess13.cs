using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    /// <summary>
    /// 尊親屬資料
    /// </summary>
    public class CheckProcess13:ICheckProcess
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
            #region RELATIVE
            List<string> chkItems1 = new List<string>();
            chkItems1.Add("直系血親_工作機構");
            chkItems1.Add("直系血親_出生年");
            chkItems1.Add("直系血親_存、歿");
            chkItems1.Add("直系血親_姓名");
            chkItems1.Add("直系血親_原國籍");
            chkItems1.Add("直系血親_教育程度");
            chkItems1.Add("直系血親_電話");
            chkItems1.Add("直系血親_稱謂");
            chkItems1.Add("直系血親_職業");
            chkItems1.Add("直系血親_職稱");
            chkItems1.Add("直系血親_行動電話");

            

            _ErrorCount += CheckDataTransfer.CheckRELATIVE_Error("家庭狀況", chkItems1, _Student);

            _TotalCount = chkItems1.Count;

            //// 這算一項
            //if (CheckDataTransfer.CheckRELATIVE_Error(_GroupName, chkItems1, _Student)>0)
            //    _ErrorCount += 1; ;
            
            //_TotalCount += 1;
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
