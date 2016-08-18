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

                foreach (string ss in Items)
                {
                    string key1 = GroupName + "_" + ss;
                    if (chkDict.ContainsKey(key1))
                    {
                        if (chkDict[key1] == null)
                            retError++;
                        else
                        { 
                            // 檢查 name 是否輸入
                            if (chkDict[key1]["name"] == null)
                                retError++;
                            else
                                if (chkDict[key1]["name"].ToString() == "")
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

                foreach (string ss in Items)
                {
                    string key1 = GroupName + "_" + ss;
                    if (chkDict.ContainsKey(key1))
                    {
                        if (chkDict[key1] == null)
                            retError++;
                        else
                        {
                            // 檢查 name 是否輸入
                            if (chkDict[key1]["name"] == null)
                                retError++;
                            else
                                if (chkDict[key1]["name"].ToString() == "")
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
    }
}
