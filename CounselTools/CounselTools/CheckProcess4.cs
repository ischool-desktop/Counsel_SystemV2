using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    /// <summary>
    /// 自傳
    /// </summary>
    public class CheckProcess4:ICheckProcess
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
            chkItems1.Add("小學（國中）老師或同學常說我是");
            chkItems1.Add("小學（國中）時我曾在班上登任過的職務有");
            chkItems1.Add("他是怎樣的人");

            List<string> items = new List<string>();
            items.Add("年級_我目前最需要的協助是");
            items.Add("年級_我目前遇到最大的困難是");
            string strItem = "";
            for (int g = 1; g <= _Student.GradeYear; g++)
            {
                foreach (string str in items)
                {
                    string ssItem = "";
                    switch (g)
                    {
                        case 1: ssItem = "一"; break;
                        case 2: ssItem = "二"; break;
                        case 3: ssItem = "三"; break;                        
                    }
                    strItem = "自我的心聲_" + ssItem + str;
                    chkItems1.Add(strItem);
                }
            }

            chkItems1.Add("我在小學（國中）得過的獎有");
            chkItems1.Add("我在家中最怕的人是");
            chkItems1.Add("我在家中最怕的人是_因為");
            chkItems1.Add("我排遣休閒時間的方法是");
            chkItems1.Add("我最難忘的一件事是");
            chkItems1.Add("我覺得我自己的過去最滿意的是");
            chkItems1.Add("我覺得我的缺點是");
            chkItems1.Add("我覺得我的優點是");
//            chkItems1.Add("我讀過且印象最深刻的課外書是");
            chkItems1.Add("家中最了解我的人");
            chkItems1.Add("家中最了解我的人_因為");
            chkItems1.Add("國中時的學校生活");
            chkItems1.Add("常指導我做功課的人");
            chkItems1.Add("喜歡的人");
            chkItems1.Add("喜歡的人_因為");
            chkItems1.Add("最不喜歡做的事");
            chkItems1.Add("最不喜歡做的事_因為");
            chkItems1.Add("最快樂的回憶");
            chkItems1.Add("最足以描述自己的幾句話");
            chkItems1.Add("最要好的朋友");
            chkItems1.Add("最喜歡的國小（國中）老師");
            chkItems1.Add("最喜歡的國小（國中）老師__因為");
            chkItems1.Add("最喜歡做的事");
            chkItems1.Add("最喜歡做的事_因為");
            chkItems1.Add("最痛苦的回憶");
//            chkItems1.Add("填寫日期");
            chkItems1.Add("讀過且印象最深刻的課外書");

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
