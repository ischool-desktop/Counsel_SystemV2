using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace CounselTools
{
    /// <summary>
    /// 檢查資料
    /// </summary>
    public class CheckDataTransfer
    {
        /// <summary>
        /// SINGLE_ANSWER
        /// </summary>
        /// <param name="GroupName"></param>
        /// <param name="Items"></param>
        /// <param name="StudentID"></param>
        /// <returns></returns>
        public static int CheckSINGLE_ANSWER_Error(string GroupName,List<string> Items, ClassStudent Student)
        {
            int retError = 0;
            if (Gobal._single_recordDict.ContainsKey(Student.StudentID))
            {
                Dictionary<string, string> chkDict = new Dictionary<string, string>();

                foreach (DataRow dr in Gobal._single_recordDict[Student.StudentID])
                {
                    string key = dr["key"].ToString();
                    if (!chkDict.ContainsKey(key))
                        chkDict.Add(key, dr["data"].ToString().Trim());
                }

                foreach (string ss in Items)
                {
                    string key = GroupName + "_" + ss;


                    //2016/9/10 穎驊註解，由於卡在目前系統既有架構之下，在最小修改原則，只能先將就這樣驗證。
                    //狀況是這樣，如果問卷填答使用者在Web2填寫輔導問卷的兄弟姊妹資料，選擇我是獨子，目前並沒有直接存這個訊息的功能
                    //取而代之則是儲存像是[家庭狀況_兄弟姊妹_排行,""]的狀況，也就是沒有兄弟姊妹排名 = 獨子
                    //而該如何判斷使用者搞不好是連填寫都沒有填寫的狀況? 與恩正、俊傑看過Web的程式碼確認後，
                    //如果是沒有填寫的狀況，此時連"家庭狀況_兄弟姊妹_排行" 此項的Key值都不會出現，所以就可以判斷

                    #region 驗證兄弟姊妹資料使用
                    if (key == "家庭狀況_兄弟姊妹_排行")
                    {
                        if (chkDict.ContainsKey(key))
                        {
                            //[家庭狀況_兄弟姊妹_排行,""]的狀況，表示"我是獨子"，先暫時回傳retError = 99999，供前一個涵式判斷;
                            if (chkDict[key] == "")
                                retError = 99999;
                        }
                        else
                            retError++;
                    } 
                    #endregion

                    else
                    {
                        if (chkDict.ContainsKey(key))
                        {
                            if (chkDict[key] == "")
                                retError++;
                        }
                        else
                            retError++;
                    }
                }
            }
            else
            {
                retError = Items.Count;
            }

            return retError;
        }

        public static int CheckMULTI_ANSWER_Error(string GroupName, List<string> Items, ClassStudent Student)
        {
            int retError = 0;
            if (Gobal._multiple_recordDict.ContainsKey(Student.StudentID))
            {
                Dictionary<string, string> chkDict = new Dictionary<string, string>();

                foreach (DataRow dr in Gobal._multiple_recordDict[Student.StudentID])
                {
                    string key = dr["key"].ToString();
                    if (!chkDict.ContainsKey(key))
                        chkDict.Add(key, dr["data"].ToString().Trim());
                }

                foreach (string ss in Items)
                {
                    string key = GroupName + "_" + ss;
                    if (chkDict.ContainsKey(key))
                    {
                        if (chkDict[key] == "")
                            retError++;
                    }
                    else
                        retError++;
                }
            }
            else
            {
                retError = Items.Count;
            }

            return retError;
        }

        public static int CheckSEMESTER_Error(string GroupName, List<string> Items, ClassStudent Student)
        {
            int retError = 0;
            string sem = K12.Data.School.DefaultSemester;
            if (Gobal._semester_dataDict.ContainsKey(Student.StudentID))
            {
                Dictionary<string, DataRow> chkDict = new Dictionary<string, DataRow>();
                foreach (DataRow dr in Gobal._semester_dataDict[Student.StudentID])
                {
                    string key = dr["key"].ToString();
                    if (!chkDict.ContainsKey(key))
                        chkDict.Add(key, dr);
                }

                foreach (string ss in Items)
                {
                    string key1 = GroupName + "_" + ss;
                    if (chkDict.ContainsKey(key1))
                    {
                        if (Student.GradeYear == 0)
                            retError++;
                        else
                        {
                            if (chkDict[key1] == null)
                                retError++;
                            else
                            {
                                bool err = false;
                                // 年級學期判斷
                                if (Student.GradeYear == 0)
                                    err = true;
                                else
                                {
                                    for (int g = 1; g <= Student.GradeYear; g++)
                                    {
                                        string kk = "s" + g + "a";
                                        string kkb = "s" + g + "b";

                                        // 只有上學期
                                        if (g == Student.GradeYear && sem == "1")
                                            kkb = kk;

                                        if (chkDict[key1][kk] == null || chkDict[key1][kkb] == null)
                                            err = true;
                                        else
                                        {
                                            if (chkDict[key1][kk].ToString() == "" || chkDict[key1][kkb].ToString() == "")
                                            {
                                                err = true;
                                            }
                                        }
                                    }
                                }
                                if (err)
                                    retError++;
                            }

                        }
                    }
                    else
                        retError++;
                }
                
            }
            else
            {
                retError = Items.Count;
            }
            return retError;     
        }

        public static int CheckYEARLY_Error(string GroupName, List<string> Items, ClassStudent Student)
        {
            int retError = 0;
            if (Gobal._yearly_dataDict.ContainsKey(Student.StudentID))
            {
                Dictionary<string, DataRow> chkDict = new Dictionary<string, DataRow>();
                foreach (DataRow dr in Gobal._yearly_dataDict[Student.StudentID])
                {
                    string key = dr["key"].ToString();
                    if (!chkDict.ContainsKey(key))
                        chkDict.Add(key, dr);
                }

                foreach (string ss in Items)
                {
                    string key1 = GroupName + "_" + ss;
                    if (chkDict.ContainsKey(key1))
                    {
                        if (chkDict[key1] == null)
                            retError++;
                        else
                        {
                            bool err = false;
                            // 年級判斷
                            if (Student.GradeYear == 0)
                                err = true;
                            else
                            {
                                for (int g = 1; g <= Student.GradeYear; g++)
                                {
                                    string kk = "g" + g;
                                    if (chkDict[key1][kk] == null)
                                        err = true;
                                    else
                                    {
                                        if (chkDict[key1][kk].ToString() == "")
                                        {
                                            err = true;
                                        }
                                    }
                                }
                            }
                            if (err)
                                retError++;
                        }
                    }
                    else
                        retError++;
                }

            }
            else
            {
                retError = Items.Count;
            }
            return retError;
        }

        public static int CheckPRIORITY_Error(string GroupName, List<string> Items, ClassStudent Student)
        {
            int retError = 0;
            if (Gobal._priority_dataDict.ContainsKey(Student.StudentID))
            {
                Dictionary<string, DataRow> chkDict = new Dictionary<string, DataRow>();
                foreach (DataRow dr in Gobal._priority_dataDict[Student.StudentID])
                {
                    string key = dr["key"].ToString();
                    if (!chkDict.ContainsKey(key))
                        chkDict.Add(key, dr);
                }

                foreach (string ss in Items)
                {
                    string key1 = GroupName + "_" + ss;
                    if (chkDict.ContainsKey(key1))
                    {
                        if (chkDict[key1] == null)
                            retError++;
                        else
                        { 
                            // 檢查優先第一項是否有輸入
                            if (chkDict[key1]["p1"] == null)
                                retError++;
                            else
                                if (chkDict[key1]["p1"].ToString() == "")
                                    retError++;
                        }
                    }
                    else
                        retError++;
                }

            }
            else
            {
                retError = Items.Count;
            }
            return retError;
        }

        public static int CheckRELATIVE_Error(string GroupName, List<string> Items, ClassStudent Student)
        {
            int retError = 0;
            if (Gobal._relativeDict.ContainsKey(Student.StudentID))
            {
                Dictionary<string, DataRow> chkDict = new Dictionary<string, DataRow>();
                foreach (DataRow dr in Gobal._relativeDict[Student.StudentID])
                {
                    string key = dr["title"].ToString();
                    if (!chkDict.ContainsKey(key))
                        chkDict.Add(key, dr);
                }



                //2016/9/9 穎驊新增，由於舊的邏輯無法處理新的直系血親關係，寫新的邏輯來做判斷未填寫數量
                List<string> chkTagNames_Relative = new List<string>();
                
                chkTagNames_Relative.Add("Name");
                chkTagNames_Relative.Add("birth_year");
                chkTagNames_Relative.Add("is_alive");
                chkTagNames_Relative.Add("Phone");
                chkTagNames_Relative.Add("Job");
                chkTagNames_Relative.Add("Institute");
                chkTagNames_Relative.Add("job_title");
                chkTagNames_Relative.Add("edu_degree");
                chkTagNames_Relative.Add("National");
                chkTagNames_Relative.Add("cell_phone");

               
               //foreach(string key_Titles in chkDict.Keys)
               // {
               //  foreach (string TagName in chkTagNames_Relative)                  
               //  {
               //      if (chkDict[key_Titles][TagName] == null || chkDict[key_Titles][TagName]+"" =="") 
               //      {
               //          retError++;                     
               //      }                    
                    
               //  }                                                
               // }

               foreach (string key_Titles in chkDict.Keys)
               {
                   foreach (string TagName in chkTagNames_Relative)
                   {
                       if (chkDict[key_Titles][TagName] == null )
                       {
                           retError++;
                       }

                   }
               }


                //foreach (string ss in Items)
                //{
                //    string key1 = GroupName + "_" + ss;
                //    if (chkDict.ContainsKey(key1))
                //    {
                //        if (chkDict[key1] == null)
                //            retError++;
                //        else
                //        { 
                //            // 檢查 name 是否輸入
                //            if (chkDict[key1]["name"] == null)
                //                retError++;
                //            else
                //                if (chkDict[key1]["name"].ToString() == "")
                //                    retError++;
                //        }
                //    }
                //    else
                //        retError++;
                //}

            }
            else
            {
                retError = Items.Count;
            }
            return retError;
        }

        public static int CheckSIBLING_Error(string GroupName, List<string> Items, ClassStudent Student)
        {
            int retError = 0;
            if (Gobal._siblingDict.ContainsKey(Student.StudentID))
            {
                Dictionary<string, DataRow> chkDict = new Dictionary<string, DataRow>();
                foreach (DataRow dr in Gobal._siblingDict[Student.StudentID])
                {
                    string key = dr["title"].ToString();
                    if (!chkDict.ContainsKey(key))
                        chkDict.Add(key, dr);
                }

                //2016/9/9 穎驊新增，由於舊的邏輯無法處理新的兄弟姊妹資料，寫新的邏輯來做判斷未填寫數量
                List<string> chkTagNames_Sibling = new List<string>();

                chkTagNames_Sibling.Add("Name");
                chkTagNames_Sibling.Add("birth_year");
                chkTagNames_Sibling.Add("school_name");
            

                foreach (string key_Titles in chkDict.Keys)
                {
                    foreach (string TagName in chkTagNames_Sibling)
                    {
                        if (chkDict[key_Titles][TagName] == null || chkDict[key_Titles][TagName] + "" == "")
                        {
                            retError++;
                        }

                    }
                }

                //foreach (string ss in Items)
                //{
                //    string key1 = GroupName + "_" + ss;
                //    if (chkDict.ContainsKey(key1))
                //    {
                //        if (chkDict[key1] == null)
                //            retError++;
                //        else
                //        {
                //            // 檢查 name 是否輸入
                //            if (chkDict[key1]["name"] == null)
                //                retError++;
                //            else
                //                if (chkDict[key1]["name"].ToString() == "")
                //                    retError++;
                //        }
                //    }
                //    else
                //        retError++;
                //}

            }
            else
            {
                retError = Items.Count;
            }
            return retError;
        }
    }
}
