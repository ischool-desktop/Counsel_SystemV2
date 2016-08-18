using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;
using K12.Data;
using System.Xml.Linq;

namespace Counsel_System2.DAO
{    
    /// <summary>
    /// 學生資料處理
    /// </summary>
    public class StudentInfoTransfer
    {
        List<string> _StudentIDList;
        List<StudentInfo> _StudentInfoList;
        List<UDT_StudQuizDataDef> _StudQuizDataDefList;
        UDTTransfer _UDTTransfer;
        List<UpdateRecordRecord> _updateRecList;
        public static Dictionary<string, Dictionary<string, UDT_StudQuizDataDef>> StudQuizDataDefDict = new Dictionary<string, Dictionary<string, UDT_StudQuizDataDef>>();
        List<UDT_QuizDef> _QuizDef;
        // 學期歷程學期年級索引
        Dictionary<string, List<DAO.SchoolYearSemester>> _StudentHistoryDict;

        public StudentInfoTransfer()
        {
            _StudentInfoList = new List<StudentInfo>();
            _StudentHistoryDict = new Dictionary<string, List<SchoolYearSemester>>();
            _updateRecList = new List<UpdateRecordRecord>();
            _StudQuizDataDefList = new List<UDT_StudQuizDataDef>();
            _UDTTransfer = new UDTTransfer();
            _QuizDef = new List<UDT_QuizDef>();
        }

        /// <summary>
        /// 取得學生資訊
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public List<StudentInfo> GetStudentInfoList(List<string> StudentIDList)
        {
            _StudentIDList = StudentIDList;

            ConvertStudentHistoryDict();
            ConvertBaseStudentInfo();
            ConvertUpdateRecord();
            ConvertAttendance();
            ConvertMerit();
            ConvertStudQuizDataDef();
            return _StudentInfoList;
        }

        /// <summary>
        /// 轉換基本資料
        /// </summary>
        private void ConvertBaseStudentInfo()
        {
            QueryHelper helper = new QueryHelper();
            if(_StudentIDList.Count>0)
            {                
                string strQuery = "select id,name,student_number,id_number,birthdate,permanent_address,mailing_address,permanent_phone,contact_phone,gender,birth_place from student where id in("+ string.Join(",",_StudentIDList.ToArray()) + ")";
                DataTable dt = helper.Select(strQuery);

                foreach (DataRow dr in dt.Rows)
                {
                    StudentInfo si = new StudentInfo();
                    si.MeritDict = new Dictionary<string, string>();
                    si.AttendanceDict = new Dictionary<string, string>();
                    si.UpdateRecordList = new List<string>();
                    si.StudentID = dr[0].ToString();
                    si.Name = dr[1].ToString();
                    si.StudentNumber = dr[2].ToString();
                    si.IDNumber = dr[3].ToString();
                    DateTime dtBirth;
                    if (DateTime.TryParse(dr[4].ToString(), out dtBirth))
                    {
                        si.Birthday = dtBirth.ToShortDateString();
                        si.BirthdayTW = "民國" + (dtBirth.Year - 1911) + "年" + dtBirth.Month + "月" + dtBirth.Day + "日";
                    }
                    si.PermanentAddress = XmlAddressConvert(dr[5].ToString());
                    si.MailingAddress=XmlAddressConvert(dr[6].ToString ());
                    si.PermanentPhone = dr[7].ToString();
                    si.MailingPhone = dr[8].ToString();
                    si.Gender = "";
                    if (dr[9].ToString() == "1")
                        si.Gender = "男";

                    if (dr[9].ToString() == "0")
                        si.Gender = "女";
                    si.Birthplace = dr[10].ToString();

                    si.SchoolName=School.ChineseName;
                    _StudentInfoList.Add(si);
                }
            }
        }

        /// <summary>
        /// 轉換異動
        /// </summary>
        private void ConvertUpdateRecord()
        {
            _updateRecList = UpdateRecord.SelectByStudentIDs(_StudentIDList);
            foreach (StudentInfo si in _StudentInfoList)
            {
                foreach (UpdateRecordRecord updateRec in _updateRecList.Where(x => x.StudentID == si.StudentID))
                    si.UpdateRecordList.Add(updateRec.UpdateDate + "," + updateRec.UpdateDescription);               
                
            }           
        
        }

        /// <summary>
        /// 轉換成完成地址
        /// </summary>
        /// <param name="xmlString"></param>
        /// <returns></returns>
        private string XmlAddressConvert(string xmlString)
        {
            StringBuilder sb = new StringBuilder ();
            if (!string.IsNullOrEmpty(xmlString))
            {
                XElement elm = XElement.Parse(xmlString);
                if (elm != null)
                {
                    if (elm.Element("Address").Element("ZipCode") != null)
                        sb.Append(elm.Element("Address").Element("ZipCode").Value);
                    if (elm.Element("Address").Element("County") != null)
                        sb.Append(elm.Element("Address").Element("County").Value);
                    if (elm.Element("Address").Element("Town") != null)
                        sb.Append(elm.Element("Address").Element("Town").Value);
                    if (elm.Element("Address").Element("District") != null)
                        sb.Append(elm.Element("Address").Element("District").Value);
                    if (elm.Element("Address").Element("Area") != null)
                        sb.Append(elm.Element("Address").Element("Area").Value);
                    if (elm.Element("Address").Element("DetailAddress") != null)
                        sb.Append(elm.Element("Address").Element("DetailAddress").Value);                
                }
            }

            return sb.ToString();
        }

        public void ConvertStudQuizDataDef()
        {
            StudQuizDataDefDict.Clear();

            _QuizDef = _UDTTransfer.GetAllQuizData();

            Dictionary<int, string> qNameDict = new Dictionary<int, string>();
            foreach (UDT_QuizDef qd in _QuizDef)
            {
                int id=int.Parse(qd.UID);
                qNameDict.Add(id, qd.QuizName);
            }

            _StudQuizDataDefList = _UDTTransfer.GetStudQuizDataByStudentIDList(_StudentIDList);
            foreach (StudentInfo si in _StudentInfoList)
            {
                if (!StudQuizDataDefDict.ContainsKey(si.StudentID))
                    StudQuizDataDefDict.Add(si.StudentID, new Dictionary<string, UDT_StudQuizDataDef>());

                int id = int.Parse(si.StudentID);
                foreach (UDT_StudQuizDataDef data in _StudQuizDataDefList.Where(x => x.StudentID == id))
                {
                    string qName = "";
                    if (qNameDict.ContainsKey(data.QuizID))
                        qName = qNameDict[data.QuizID];

                    if (qName != "")
                    {
                        if (!StudQuizDataDefDict[si.StudentID].ContainsKey(qName))
                            StudQuizDataDefDict[si.StudentID].Add(qName, data);
                    }
                }
            }
        }

        /// <summary>
        /// 轉換缺曠
        /// </summary>
        private void ConvertAttendance()
        {
            Dictionary<string, Dictionary<string, int>> countDict = new Dictionary<string, Dictionary<string, int>>();
            // 讀取缺曠資料            
            QueryHelper helper = new QueryHelper();
            if (_StudentIDList.Count > 0)
            {
                string strQuery = "select ref_student_id,school_year,semester,detail from attendance where ref_student_id in("+string.Join(",",_StudentIDList.ToArray())+")";
                DataTable dt = helper.Select(strQuery);
                foreach (DataRow dr in dt.Rows)
                {
                    // StudentID_SchoolYear_Semester
                    string key = dr[0].ToString()+"_"+dr[1].ToString()+"_"+ dr[2].ToString();
                    // Detail
                    XElement detail = XElement.Parse(dr[3].ToString());

                    if (!countDict.ContainsKey(key))
                        countDict.Add(key, new Dictionary<string, int>());

                    if(detail !=null)
                    {
                        foreach(XElement elm in detail.Elements("Period"))
                        {
                            string val=elm.Attribute("AbsenceType").Value;
                            if (countDict[key].ContainsKey(val))
                                countDict[key][val]++;
                            else
                                countDict[key].Add(val, 1);                            
                        }
                    }     
                }

                // 解析
                foreach (StudentInfo si in _StudentInfoList)
                { 
                    if(_StudentHistoryDict.ContainsKey(si.StudentID))
                    {
                        foreach (SchoolYearSemester sys in _StudentHistoryDict[si.StudentID])
                        {
                            string key1 = si.StudentID + "_" + sys.SchoolYear + "_" + sys.Semester;
                            if (countDict.ContainsKey(key1))
                            {
                                string str = "";
                                if (sys.GradeYear == "1")
                                    str = "一";
                                if (sys.GradeYear == "2")
                                    str = "二";
                                if (sys.GradeYear == "3")
                                    str = "三";
                                if (sys.GradeYear == "4")
                                    str = "四";
                                if (sys.GradeYear == "7")
                                    str = "七";
                                if (sys.GradeYear == "8")
                                    str = "八";
                                if (sys.GradeYear == "9")
                                    str = "九";

                                if (sys.Semester == "1")
                                    str += "上";

                                if (sys.Semester == "2")
                                    str += "下";

                                
                                StringBuilder sb = new StringBuilder ();
                                foreach(KeyValuePair<string,int> val  in countDict[key1])
                                    sb.Append(val.Key+":"+val.Value+";");

                                if (!si.AttendanceDict.ContainsKey(str))
                                    si.AttendanceDict.Add(str, sb.ToString());
                            }                            
                        }
                    }                    
                }
            }
         
        }

        /// <summary>
        /// 轉換獎勵
        /// </summary>
        private void ConvertMerit()
        {
            Dictionary<string, Dictionary<string, int>> countDict = new Dictionary<string, Dictionary<string, int>>();
            // 讀取獎懲資料            
            QueryHelper helper = new QueryHelper();
            if (_StudentIDList.Count > 0)
            {
                string strQuery = "select ref_student_id,school_year,semester,detail,type from discipline where ref_student_id in(" + string.Join(",", _StudentIDList.ToArray()) + ")";
                DataTable dt = helper.Select(strQuery);
                foreach (DataRow dr in dt.Rows)
                {
                    // StudentID_SchoolYear_Semester
                    string key = dr[0].ToString() + "_" + dr[1].ToString() + "_" + dr[2].ToString();
                    // Detail
                    XElement detail = XElement.Parse(dr[3].ToString());

                    if (!countDict.ContainsKey(key))
                        countDict.Add(key, new Dictionary<string, int>());

                    if (detail != null)
                    {
                        // 功
                        if(detail.Element("Merit") !=null)
                        {
                           int iA,iB,iC;
                           if (int.TryParse(detail.Element("Merit").Attribute("A").Value, out iA))
                               if (iA > 0)
                               {
                                   if (countDict[key].ContainsKey("大功"))
                                       countDict[key]["大功"] += iA;
                                   else
                                       countDict[key].Add("大功", iA);
                               }

                           if (int.TryParse(detail.Element("Merit").Attribute("B").Value, out iB))
                               if (iB > 0)
                               {
                                   if (countDict[key].ContainsKey("小功"))
                                       countDict[key]["小功"] += iB;
                                   else
                                       countDict[key].Add("小功", iB);
                               }

                           if (int.TryParse(detail.Element("Merit").Attribute("C").Value, out iC))
                               if (iC > 0)
                               {
                                   if (countDict[key].ContainsKey("嘉獎"))
                                       countDict[key]["嘉獎"] += iC;
                                   else
                                       countDict[key].Add("嘉獎", iC);
                               }
                        }
                        

                        // 過
                        if (detail.Element("Demerit") != null)
                        {
                            // 當有消過不列入
                            if (detail.Element("Demerit").Attribute("Cleared").Value!="是")
                            {

                                int iA, iB, iC;
                                if (int.TryParse(detail.Element("Demerit").Attribute("A").Value, out iA))
                                    if (iA > 0)
                                    {
                                        if (countDict[key].ContainsKey("大過"))
                                            countDict[key]["大過"] += iA;
                                        else
                                            countDict[key].Add("大過", iA);
                                    }

                                if (int.TryParse(detail.Element("Demerit").Attribute("B").Value, out iB))
                                    if (iB > 0)
                                    {
                                        if (countDict[key].ContainsKey("小過"))
                                            countDict[key]["小過"] += iB;
                                        else
                                            countDict[key].Add("小過", iB);
                                    }

                                if (int.TryParse(detail.Element("Demerit").Attribute("C").Value, out iC))
                                    if (iC > 0)
                                    {
                                        if (countDict[key].ContainsKey("警告"))
                                            countDict[key]["警告"] += iC;
                                        else
                                            countDict[key].Add("警告", iC);
                                    }
                            }
                        }
                    }
                }

                // 解析
                foreach (StudentInfo si in _StudentInfoList)
                {
                    if (_StudentHistoryDict.ContainsKey(si.StudentID))
                    {
                        foreach (SchoolYearSemester sys in _StudentHistoryDict[si.StudentID])
                        {
                            string key1 = si.StudentID + "_" + sys.SchoolYear + "_" + sys.Semester;
                            if (countDict.ContainsKey(key1))
                            {
                                string str = "";
                                if (sys.GradeYear == "1")
                                    str = "一";
                                if (sys.GradeYear == "2")
                                    str = "二";
                                if (sys.GradeYear == "3")
                                    str = "三";
                                if (sys.GradeYear == "4")
                                    str = "四";
                                if (sys.GradeYear == "7")
                                    str = "七";
                                if (sys.GradeYear == "8")
                                    str = "八";
                                if (sys.GradeYear == "9")
                                    str = "九";
                                
                                if (sys.Semester == "1")
                                    str += "上";

                                if (sys.Semester == "2")
                                    str += "下";

                               
                                StringBuilder sb = new StringBuilder();
                                foreach (KeyValuePair<string, int> val in countDict[key1])
                                    sb.Append(val.Key + ":" + val.Value + ";");

                                if (!si.MeritDict.ContainsKey(str))
                                    si.MeritDict.Add(str, sb.ToString());
                            }
                        }
                    }
                }
            }        
        }

        /// <summary>
        /// 轉換學期歷程
        /// </summary>
        private void ConvertStudentHistoryDict()
        {
            List<SemesterHistoryRecord> semsHis = SemesterHistory.SelectByStudentIDs(_StudentIDList);
            foreach (SemesterHistoryRecord shrec in semsHis)
                foreach (SemesterHistoryItem shi in shrec.SemesterHistoryItems)
                {
                    SchoolYearSemester sys = new SchoolYearSemester();
                    sys.SchoolYear = shi.SchoolYear.ToString();
                    sys.Semester = shi.Semester.ToString();
                    sys.GradeYear = shi.GradeYear.ToString();

                    if (!_StudentHistoryDict.ContainsKey(shi.RefStudentID))
                        _StudentHistoryDict.Add(shi.RefStudentID, new List<SchoolYearSemester>());

                    _StudentHistoryDict[shi.RefStudentID].Add(sys);
                }
        }


    }
}
