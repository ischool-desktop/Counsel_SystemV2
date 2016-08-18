using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    /// <summary>
    /// 家庭狀況
    /// </summary>
    public class CheckProcess2:ICheckProcess
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
            // 這算一項
            if (CheckDataTransfer.CheckRELATIVE_Error(_GroupName, chkItems1, _Student)>0)
                _ErrorCount += 1; ;
            
            _TotalCount += 1;
            #endregion
            
            #region SIBLING
            List<string> chkItems2 = new List<string>();
            chkItems2.Add("兄弟姊妹_出生年次");
            chkItems2.Add("兄弟姊妹_姓名");
            chkItems2.Add("兄弟姊妹_畢肆業學校");
            chkItems2.Add("兄弟姊妹_備註");
            chkItems2.Add("兄弟姊妹_稱謂");

            // 這算一項
            if (CheckDataTransfer.CheckSIBLING_Error(_GroupName, chkItems2, _Student) > 0)
                _ErrorCount += 1;
            
            _TotalCount += 1;

            #endregion

            #region SINGLE_ANSWER

            List<string> chkItems3 = new List<string>();
            chkItems3.Add("兄弟姊妹_排行");
            chkItems3.Add("監護人_姓名");
            chkItems3.Add("監護人_性別");
            chkItems3.Add("監護人_通訊地址");
            chkItems3.Add("監護人_電話");
            chkItems3.Add("監護人_關係");
            
            // 這算一項
            if (CheckDataTransfer.CheckSINGLE_ANSWER_Error(_GroupName, chkItems3, _Student)>0)
                _ErrorCount += 1;

            _TotalCount += 1;

            #endregion

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

            // 這算一項
            if (CheckDataTransfer.CheckYEARLY_Error(_GroupName, chkItems4, _Student) > 0)
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
