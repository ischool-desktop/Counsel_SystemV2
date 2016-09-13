using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace CounselTools
{
    /// <summary>
    /// 兄弟姊妹資料
    /// </summary>
    public class CheckProcess15:ICheckProcess
    {
        string _GroupName;
        ClassStudent _Student;
        int _ErrorCount = 0, _TotalCount = 0;
        Dictionary<string, string> _ErrorDict = new Dictionary<string, string>();


        Boolean IamTheOnlySon =false;

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

            List<string> chkItems3 = new List<string>();
            chkItems3.Add("兄弟姊妹_排行");
            //chkItems3.Add("監護人_姓名");
            //chkItems3.Add("監護人_性別");
            //chkItems3.Add("監護人_通訊地址");
            //chkItems3.Add("監護人_電話");
            //chkItems3.Add("監護人_關係");

            _ErrorCount += CheckDataTransfer.CheckSINGLE_ANSWER_Error("家庭狀況", chkItems3, _Student);

            _TotalCount += chkItems3.Count;


            //2016/9/10 穎驊註解，如果這邊看不懂可以去看CheckDataTransfer.CheckSINGLE_ANSWER_Error涵式的註解說明，
            //基本上礙於現有架構之下，只能先暫時這樣子做，如果回傳 _ErrorCount == 99999，則我們視此資料為獨子，就不會再做兄弟姊妹的判斷

            if (_ErrorCount == 99999) 
            {
                IamTheOnlySon = true;

                _ErrorCount = 0;
            }


            //// 這算一項
            //if (CheckDataTransfer.CheckSINGLE_ANSWER_Error("家庭狀況", chkItems3, _Student) > 0)
            //    _ErrorCount += 1;

            //_TotalCount += 1;

            #endregion       

            //我是獨子=false 才會繼續
            if (IamTheOnlySon == false)
            {
                #region SIBLING
                List<string> chkItems2 = new List<string>();
                chkItems2.Add("兄弟姊妹_出生年次");
                chkItems2.Add("兄弟姊妹_姓名");
                chkItems2.Add("兄弟姊妹_畢肆業學校");                
                chkItems2.Add("兄弟姊妹_稱謂");
                //chkItems2.Add("兄弟姊妹_備註");

                _ErrorCount += CheckDataTransfer.CheckSIBLING_Error("家庭狀況", chkItems2, _Student);
                _TotalCount += chkItems2.Count;

                //// 這算一項
                //if (CheckDataTransfer.CheckSIBLING_Error("家庭狀況", chkItems2, _Student) > 0)
                //    _ErrorCount += 1;

                //_TotalCount += 1;

                #endregion
            }


            //#region YEARLY
            //List<string> chkItems4 = new List<string>();
            //chkItems4.Add("父母關係");
            //chkItems4.Add("父親管教方式");
            //chkItems4.Add("本人住宿");
            //chkItems4.Add("母親管教方式");
            //chkItems4.Add("我覺得是否足夠");
            //chkItems4.Add("每星期零用錢");
            //chkItems4.Add("居住環境");
            //chkItems4.Add("家庭氣氛");
            //chkItems4.Add("經濟狀況");

            //// 這算一項
            //if (CheckDataTransfer.CheckYEARLY_Error(_GroupName, chkItems4, _Student) > 0)
            //    _ErrorCount += 1;
            //_TotalCount += 1;

            //#endregion
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
