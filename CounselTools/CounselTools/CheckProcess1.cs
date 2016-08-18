using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    /// <summary>
    /// 檢查本人概況資料
    /// </summary>
    public class CheckProcess1:ICheckProcess
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
            // 本人概況	生理缺陷	MULTI_ANSWER
            // 本人概況	血型	SINGLE_ANSWER
            // 本人概況	身高	SEMESTER
            // 本人概況	宗教	SINGLE_ANSWER
            // 本人概況	原住民血統	SINGLE_ANSWER
            // 本人概況	曾患特殊疾病	MULTI_ANSWER
            // 本人概況	體重	SEMESTER


            #region SINGLE_ANSWER
            List<string> chkItems1 = new List<string>();
            chkItems1.Add("血型");
            chkItems1.Add("宗教");
            chkItems1.Add("原住民血統");

            // SINGLE_ANSWER
            _ErrorCount += CheckDataTransfer.CheckSINGLE_ANSWER_Error(_GroupName, chkItems1, _Student);
            _TotalCount += chkItems1.Count;
            #endregion

            #region MULTI_ANSWER
            List<string> chkItems2 = new List<string>();
            chkItems2.Add("生理缺陷");
            chkItems2.Add("曾患特殊疾病");
            _ErrorCount += CheckDataTransfer.CheckMULTI_ANSWER_Error(_GroupName, chkItems2, _Student);
            _TotalCount += chkItems2.Count;

            #endregion

            #region SEMESTER
            // 這算一項
            List<string> chkItems3 = new List<string>();
            chkItems3.Add("身高");
            chkItems3.Add("體重");

            if (CheckDataTransfer.CheckSEMESTER_Error(_GroupName, chkItems3, _Student) > 0)
                _ErrorCount += 1;

            _TotalCount += 1;

            #endregion
        }


//        private void ChkDataSD(Dictionary<streing>)

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
