using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using System.Xml.Linq;
using K12.Data;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using FISCA.Data;

namespace Counsel_System2.Forms
{
    public partial class AllCounselStatisticsForm : FISCA.Presentation.Controls.BaseForm
    {
        private List<string> _StudentIDList = new List<string>();

        private string defaultYear;
        private string defaultSemester;

        //學生ID 與科別的對照
        private Dictionary<string, string> stuIdToDeptDict = new Dictionary<string, string>();

        //學生ID 與年級的對照
        private Dictionary<string, string> stuIdToGradeDict = new Dictionary<string, string>();

        // 欄數與英文的對照
        private Dictionary<int, string> colLetter = new Dictionary<int, string>();

        // 老師ID與其下所有"有晤談或是聯繫"的學生IDList
        private Dictionary<string, List<int>> teacher_studentList_Dict = new Dictionary<string, List<int>>();

        // 輔導類型與其下所有 資料
        private Dictionary<string, DAO.AllCounselCaseStaticticsRecord> counselCase_Dict = new Dictionary<string, DAO.AllCounselCaseStaticticsRecord>();


        // 輔導老師與其下所有 訪談資料
        private Dictionary<string, DAO.CounselTeacherHomeVisitWaysRecord> counselTeacherHomeVisit_Dict = new Dictionary<string, DAO.CounselTeacherHomeVisitWaysRecord>();

        // 輔導老師與其下所有 晤談資料
        private Dictionary<string, DAO.CounselTeacherInterviewWaysRecord> counselTeacherInterview_Dict = new Dictionary<string, DAO.CounselTeacherInterviewWaysRecord>();

        // 老師(高一、高二、高三)與其下所有 訪談資料
        private Dictionary<string, DAO.TeacherHomeVisitWaysRecord> TeacherHomeVisit_Dict = new Dictionary<string, DAO.TeacherHomeVisitWaysRecord>();

        // 老師(高一、高二、高三)與其下所有 晤談資料
        private Dictionary<string, DAO.CounselTeacherInterviewRecord> TeacherInterview_Dict = new Dictionary<string, DAO.CounselTeacherInterviewRecord>();


        //所有的輔導案件類別
        private List<string> counselTypeKindList = new List<string>();

        //所有的輔導老師 訪談方式
        private List<string> counselTeacherHomeVisitWaysList = new List<string>();

        //所有的輔導老師 晤談方式
        private List<string> counselTeacherInterviewWaysList = new List<string>();

        //所有的老師(高一、高二、高三) 訪談方式
        private List<string> TeacherHomeVisitWaysList = new List<string>();

        //所有的老師(高一、高二、高三) 晤談方式
        private List<string> TeacherInterviewWaysList = new List<string>();

        //所有的年級 
        private List<string> gradeYearList = new List<string>();


        public AllCounselStatisticsForm()
        {
            InitializeComponent();

            
            counselTypeKindList.Add("家人議題");
            counselTypeKindList.Add("違規行為");
            counselTypeKindList.Add("心理困擾");
            counselTypeKindList.Add("學習問題");
            counselTypeKindList.Add("性別議題");
            counselTypeKindList.Add("人際關係");
            counselTypeKindList.Add("生涯規劃");
            counselTypeKindList.Add("自傷/自殺");
            counselTypeKindList.Add("生活適應");
            counselTypeKindList.Add("生活作息/常規");
            counselTypeKindList.Add("家長期望");
            counselTypeKindList.Add("健康問題");
            counselTypeKindList.Add("情緒不穩");
            counselTypeKindList.Add("法定通報-兒少保護");
            counselTypeKindList.Add("法定通報-高風險家庭");
            counselTypeKindList.Add("法定通報-家暴(18 歲以上)");
            counselTypeKindList.Add("其他(含生活關懷)");

            //另外補上老師身分 (與輔導案例接近的分類)
            counselTypeKindList.Add("班導師");
            counselTypeKindList.Add("任輔老師");
            counselTypeKindList.Add("輔導老師");

            //將輔導案例加入字典
            foreach (string counselcase in counselTypeKindList)
            {
                counselCase_Dict.Add(counselcase, new DAO.AllCounselCaseStaticticsRecord());

                //建立 其子字典
                counselCase_Dict[counselcase].CaseStaticsPeopleCountDict = new Dictionary<string, int>();
                counselCase_Dict[counselcase].CaseStaticsPeopleDict = new Dictionary<string, List<string>>();
            }

            List<string> authorRoleList = new List<string>();
            authorRoleList.Add("輔導老師");
            authorRoleList.Add("認輔老師");
            authorRoleList.Add("班導師");

            // 學期 為 學校系統 設定的學年度
            defaultYear = K12.Data.School.DefaultSchoolYear;
            defaultSemester = K12.Data.School.DefaultSemester;


            //建立 所有的輔導老師 訪談方式 
            // 2017/1/15 穎驊註記， 訪談為 電話連絡，而非 電話 ， 先暫時寫死，未來要能夠動態支援顯示每次不同的方式
            counselTeacherHomeVisitWaysList.Add("電話聯絡");
            counselTeacherHomeVisitWaysList.Add("家庭訪問");
            counselTeacherHomeVisitWaysList.Add("家長座談");
            counselTeacherHomeVisitWaysList.Add("個別約談家長");
            counselTeacherHomeVisitWaysList.Add("其他");

            //建立 所有的輔導老師 晤談方式
            counselTeacherInterviewWaysList.Add("電話");
            counselTeacherInterviewWaysList.Add("面談");

            //建立 所有的老師老師(高一、高二、高三) 訪談方式 
            // 2017/1/15 穎驊註記， 訪談為 電話連絡，而非 電話 ， 先暫時寫死，未來要能夠動態支援顯示每次不同的方式
            TeacherHomeVisitWaysList.Add("電話聯絡");
            TeacherHomeVisitWaysList.Add("家庭訪問");
            TeacherHomeVisitWaysList.Add("家長座談");
            TeacherHomeVisitWaysList.Add("個別約談家長");
            TeacherHomeVisitWaysList.Add("其他");

            TeacherInterviewWaysList.Add("電話");
            TeacherInterviewWaysList.Add("面談");

            gradeYearList.Add("1");
            gradeYearList.Add("2");
            gradeYearList.Add("3");

            #region 建立行數 英文對照表
            // 建立行數 英文對照表
            colLetter.Add(0, "A");
            colLetter.Add(1, "B");
            colLetter.Add(2, "C");
            colLetter.Add(3, "D");
            colLetter.Add(4, "E");
            colLetter.Add(5, "F");
            colLetter.Add(6, "G");
            colLetter.Add(7, "H");
            colLetter.Add(8, "I");
            colLetter.Add(9, "J");
            colLetter.Add(10, "K");
            colLetter.Add(11, "L");
            colLetter.Add(12, "M");
            colLetter.Add(13, "N");
            colLetter.Add(14, "O");
            colLetter.Add(15, "P");
            colLetter.Add(16, "Q");
            colLetter.Add(17, "R");
            colLetter.Add(18, "S");
            colLetter.Add(19, "T");
            colLetter.Add(20, "U");
            colLetter.Add(21, "V");
            colLetter.Add(22, "W");
            colLetter.Add(23, "X");
            colLetter.Add(24, "Y");
            colLetter.Add(25, "Z"); 
            #endregion

        }



        private void buttonX1_Click(object sender, EventArgs e)
        {

            var filterSchoolyear = defaultYear;
            var filterSemester = defaultSemester;
            var filterDateBegin = dateTimeInput1.Value;
            var filterDateEnd = dateTimeInput2.Value.AddDays(1);





            Workbook wb = new Workbook();

            BackgroundWorker printWorker = new BackgroundWorker() { WorkerReportsProgress = true };
            printWorker.DoWork += delegate
            {
                FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

                // 2017/1/11 穎驊 註記 我們公司API AccessHelper 沒有 between 的方法 ，只好先用學期篩完後，再到C#操作 內 使用日期篩
                //var interviewRecordList = _AccessHelper.Select<DAO.UDT_CounselStudentInterviewRecordDef>("interview_date between '2017-09-01 00:00:00' AND '2017-12-31 23:59:59'");

                //抓晤談紀錄
                var interviewRecordList = _AccessHelper.Select<DAO.UDT_CounselStudentInterviewRecordDef>("schoolyear =" + "'" + filterSchoolyear + "'" + " and semester =" + "'" + filterSemester + "'");

                //抓家庭聯繫紀錄
                var HomeVisitRecordList = _AccessHelper.Select<DAO.UDT_Counsel_home_visit_RecordDef>("schoolyear =" + "'" + filterSchoolyear + "'" + " and semester =" + "'" + filterSemester + "'");

                //抓班級 年級、科別、導師的對應 
                QueryHelper hepler = new QueryHelper();
                string strSQL = "SELECT class_name,grade_year,class.ref_teacher_id,class.status,display_order,ref_dept_id,teacher.teacher_name,dept.name AS dept_name  FROM class INNER JOIN teacher ON  class.ref_teacher_id=teacher.id INNER JOIN dept ON class.ref_dept_id =dept.id";
                DataTable dt_class = hepler.Select(strSQL);

                //抓班級 年級、科別、學生的對應  
                strSQL = "SELECT student.id,student.name,student.status,student.ref_class_id,class_name,dept.name AS dept_name,class.grade_year FROM student INNER JOIN class ON student.ref_class_id=class.id INNER JOIN dept ON class.ref_dept_id =dept.id";
                DataTable dt_student = hepler.Select(strSQL);

                //抓系統內的輔導老師
                strSQL = "select distinct teacher.id as id1,tag_teacher.id as id2,teacher.teacher_name,teacher.nickname,tag.name,tag.prefix from teacher inner join tag_teacher on teacher.id=tag_teacher.ref_teacher_id inner join tag on tag_teacher.ref_tag_id=tag.id where tag.prefix='輔導' and tag.name='輔導老師';";
                DataTable dt_counselTeacher = hepler.Select(strSQL);

                List<DAO.AllCounselStaticticsTeacherRecord> CounselTeacherList_grade1 = new List<DAO.AllCounselStaticticsTeacherRecord>();
                List<DAO.AllCounselStaticticsTeacherRecord> CounselTeacherList_grade2 = new List<DAO.AllCounselStaticticsTeacherRecord>();
                List<DAO.AllCounselStaticticsTeacherRecord> CounselTeacherList_grade3 = new List<DAO.AllCounselStaticticsTeacherRecord>();

                //整理導師班級
                foreach (DataRow dr in dt_class.Rows)
                {
                    DAO.AllCounselStaticticsTeacherRecord AllcounselTeacherRec = new DAO.AllCounselStaticticsTeacherRecord();

                    // 班級名稱
                    AllcounselTeacherRec.ClassName = dr[0].ToString();
                    // 班級年級
                    AllcounselTeacherRec.ClassGrade = dr[1].ToString();
                    // 教師編號
                    AllcounselTeacherRec.TeacherID = dr[2].ToString();
                    // 班級狀態
                    AllcounselTeacherRec.ClassStatus = dr[3].ToString();
                    // 班級排序
                    AllcounselTeacherRec.ClassDisplayOrder = int.Parse(dr[4].ToString());

                    // 教師姓名
                    AllcounselTeacherRec.TeacherName = dr[6].ToString();
                    // 班級科別
                    AllcounselTeacherRec.ClassDepartment = dr[7].ToString();


                    if (AllcounselTeacherRec.ClassGrade == "1")
                    {
                        CounselTeacherList_grade1.Add(AllcounselTeacherRec);
                    }
                    if (AllcounselTeacherRec.ClassGrade == "2")
                    {
                        CounselTeacherList_grade2.Add(AllcounselTeacherRec);
                    }
                    if (AllcounselTeacherRec.ClassGrade == "3")
                    {
                        CounselTeacherList_grade3.Add(AllcounselTeacherRec);
                    }

                }
                // 排序
                CounselTeacherList_grade1.Sort((x, y) => { return x.ClassDisplayOrder.CompareTo(y.ClassDisplayOrder); });
                CounselTeacherList_grade2.Sort((x, y) => { return x.ClassDisplayOrder.CompareTo(y.ClassDisplayOrder); });
                CounselTeacherList_grade3.Sort((x, y) => { return x.ClassDisplayOrder.CompareTo(y.ClassDisplayOrder); });


                //整理學生科別
                foreach (DataRow dr in dt_student.Rows)
                {
                    //加入 [學生ID ，科別]
                    stuIdToDeptDict.Add(dr[0].ToString(), dr[5].ToString());

                    //加入 [學生ID ，年級]
                    stuIdToGradeDict.Add(dr[0].ToString(), dr[6].ToString());

                }


                //整理輔導老師_訪談對照
                foreach (DataRow dr in dt_counselTeacher.Rows)
                {
                    DAO.CounselTeacherHomeVisitWaysRecord record = new DAO.CounselTeacherHomeVisitWaysRecord();

                    record.CounselTeacherName = dr[2].ToString();

                    record.WaysStaticsPeopleDict = new Dictionary<string, List<string>>();

                    //依序加入 訪談方式種類
                    foreach (string ways in counselTeacherHomeVisitWaysList)
                    {
                        record.WaysStaticsPeopleDict.Add(ways, new List<string>());
                    }

                    record.WaysStaticsPeopleCountDict = new Dictionary<string, int>();

                    //依序加入 訪談方式種類
                    foreach (string ways in counselTeacherHomeVisitWaysList)
                    {
                        record.WaysStaticsPeopleCountDict.Add(ways, 0);
                    }

                    //加入 [輔導老師ID ，資料]
                    counselTeacherHomeVisit_Dict.Add(dr[0].ToString(), record);
                }


                //整理輔導老師_晤談對照
                foreach (DataRow dr in dt_counselTeacher.Rows)
                {
                    DAO.CounselTeacherInterviewWaysRecord record = new DAO.CounselTeacherInterviewWaysRecord();

                    record.CounselTeacherName = dr[2].ToString();

                    record.WaysStaticsPeopleDict = new Dictionary<string, List<string>>();

                    //依序加入 晤談方式種類
                    foreach (string ways in counselTeacherInterviewWaysList)
                    {
                        record.WaysStaticsPeopleDict.Add(ways, new List<string>());
                    }

                    record.WaysStaticsPeopleCountDict = new Dictionary<string, int>();

                    //依序加入 晤談方式種類
                    foreach (string ways in counselTeacherInterviewWaysList)
                    {
                        record.WaysStaticsPeopleCountDict.Add(ways, 0);
                    }

                    //加入 [輔導老師ID ，資料]
                    counselTeacherInterview_Dict.Add(dr[0].ToString(), record);

                    
                }


                // 整理老師(高一、高二、高三)_訪談對照
                foreach (string grade in gradeYearList)
                {
                    DAO.TeacherHomeVisitWaysRecord record = new DAO.TeacherHomeVisitWaysRecord();

                    if (grade == "1")
                    {
                        record.CounselTeacherName = "高一導";
                    }
                    if (grade == "2")
                    {
                        record.CounselTeacherName = "高二導";
                    }
                    if (grade == "3")
                    {
                        record.CounselTeacherName = "高三導";
                    }

                    record.WaysStaticsPeopleDict = new Dictionary<string, List<string>>();

                    //依序加入 訪談方式種類
                    foreach (string ways in TeacherHomeVisitWaysList)
                    {
                        record.WaysStaticsPeopleDict.Add(ways, new List<string>());
                    }

                    record.WaysStaticsPeopleCountDict = new Dictionary<string, int>();

                    //依序加入 訪談方式種類
                    foreach (string ways in TeacherHomeVisitWaysList)
                    {
                        record.WaysStaticsPeopleCountDict.Add(ways, 0);
                    }

                    //加入 [輔導老師ID ，資料]
                    TeacherHomeVisit_Dict.Add(grade, record);
                }
                
                // 整理老師(高一、高二、高三)_晤談對照
                foreach (string grade in gradeYearList)
                {
                    DAO.CounselTeacherInterviewRecord record = new DAO.CounselTeacherInterviewRecord();

                    if (grade == "1")
                    {
                        record.CounselTeacherName = "高一導";
                    }
                    if (grade == "2")
                    {
                        record.CounselTeacherName = "高二導";
                    }
                    if (grade == "3")
                    {
                        record.CounselTeacherName = "高三導";
                    }

                    record.WaysStaticsPeopleDict = new Dictionary<string, List<string>>();

                    //依序加入 訪談方式種類
                    foreach (string ways in TeacherHomeVisitWaysList)
                    {
                        record.WaysStaticsPeopleDict.Add(ways, new List<string>());
                    }

                    record.WaysStaticsPeopleCountDict = new Dictionary<string, int>();

                    //依序加入 訪談方式種類
                    foreach (string ways in TeacherHomeVisitWaysList)
                    {
                        record.WaysStaticsPeopleCountDict.Add(ways, 0);
                    }

                    //加入 [輔導老師ID ，資料]
                    TeacherInterview_Dict.Add(grade, record);
                }




                printWorker.ReportProgress(5);

                #region 篩選晤談紀錄(日期區間)

                Dictionary<string, List<DAO.UDT_CounselStudentInterviewRecordDef>> dicStudentInterviewRecord = new Dictionary<string, List<DAO.UDT_CounselStudentInterviewRecordDef>>();

                Dictionary<string, List<DAO.UDT_Counsel_home_visit_RecordDef>> dicHomeVisitRecord = new Dictionary<string, List<DAO.UDT_Counsel_home_visit_RecordDef>>();
                printWorker.ReportProgress(20);

                int progressCount = 0;
                //整理晤談資料，符合條件才加入
                foreach (var interviewRecord in interviewRecordList)
                {
                    //2017/1/20 穎驊更新，發現使用者有些晤談紀錄會沒有晤談日期，會造成interviewRecord.InterviewDate.Value 其中的 InterviewDate 變成 nullReference 的狀況爆掉
                    //因此在這邊處理，若沒有晤談資料沒有晤談日期又有勾選日期區間篩選，因為無法比較日期，直接排除掉此晤談資料
                    if (interviewRecord.InterviewDate == null)
                    {
                        continue;
                    }
                    else
                    {
                        if (interviewRecord.InterviewDate.Value < filterDateBegin ||
                            interviewRecord.InterviewDate >= filterDateEnd)
                        {
                            continue;
                        }
                    }

                    // 加入資料， Key 為 TeacherID
                    if (!dicStudentInterviewRecord.ContainsKey("" + interviewRecord.TeacherID))
                    {
                        dicStudentInterviewRecord.Add("" + interviewRecord.TeacherID, new List<DAO.UDT_CounselStudentInterviewRecordDef>());

                        dicStudentInterviewRecord["" + interviewRecord.TeacherID].Add(interviewRecord);

                    }
                    else
                    {
                        dicStudentInterviewRecord["" + interviewRecord.TeacherID].Add(interviewRecord);
                    }

                    progressCount++;

                    printWorker.ReportProgress(20 + 25 * progressCount / interviewRecordList.Count);
                }

                //整理聯繫紀錄，符合條件才加入
                foreach (var homeVisitRecord in HomeVisitRecordList)
                {

                    if (homeVisitRecord.home_visit_date.Value < filterDateBegin ||
                            homeVisitRecord.home_visit_date >= filterDateEnd)
                    {
                        continue;
                    }

                    // 加入資料， Key 為 TeacherID
                    if (!dicHomeVisitRecord.ContainsKey("" + homeVisitRecord.TeacherID))
                    {
                        dicHomeVisitRecord.Add("" + homeVisitRecord.TeacherID, new List<DAO.UDT_Counsel_home_visit_RecordDef>());

                        dicHomeVisitRecord["" + homeVisitRecord.TeacherID].Add(homeVisitRecord);

                    }
                    else
                    {
                        dicHomeVisitRecord["" + homeVisitRecord.TeacherID].Add(homeVisitRecord);
                    }

                    progressCount++;

                    printWorker.ReportProgress(45 + 25 * progressCount / HomeVisitRecordList.Count);
                }

                #endregion

                //晤談紀錄
                foreach (KeyValuePair<string, List<DAO.UDT_CounselStudentInterviewRecordDef>> record in dicStudentInterviewRecord)
                {
                    #region 導師統計(年級)
                    //一年級老師
                    foreach (DAO.AllCounselStaticticsTeacherRecord teacherRecord in CounselTeacherList_grade1)
                    {
                        if (record.Key == teacherRecord.TeacherID)
                        {
                            //人數
                            teacherRecord.CounselPeople = GetCounselPeople(record.Key, record.Value);
                            //加上人次
                            teacherRecord.CounselPeopleCount += record.Value.Count();
                        }
                    }
                    //二年級老師
                    foreach (DAO.AllCounselStaticticsTeacherRecord teacherRecord in CounselTeacherList_grade2)
                    {
                        if (record.Key == teacherRecord.TeacherID)
                        {
                            //人數
                            teacherRecord.CounselPeople = GetCounselPeople(record.Key, record.Value);
                            //加上人次
                            teacherRecord.CounselPeopleCount += record.Value.Count();
                        }
                    }
                    //三年級老師
                    foreach (DAO.AllCounselStaticticsTeacherRecord teacherRecord in CounselTeacherList_grade3)
                    {
                        if (record.Key == teacherRecord.TeacherID)
                        {
                            //人數
                            teacherRecord.CounselPeople = GetCounselPeople(record.Key, record.Value);
                            //加上人次
                            teacherRecord.CounselPeopleCount += record.Value.Count();
                        }
                    }
                    #endregion

                    #region 輔導案件類別
                    // 輔導案件類別
                    foreach (DAO.UDT_CounselStudentInterviewRecordDef interviewRecord in record.Value)
                    {

                        #region 解析輔導類型Xml
                        XmlDocument doc3 = new XmlDocument();
                        //幫忙加根目錄
                        string xmlContent3 = "<root>" + interviewRecord.CounselTypeKind + "</root>";
                        doc3.LoadXml(xmlContent3);
                        XmlNode newNode3 = doc3.DocumentElement;
                        doc3.AppendChild(newNode3);
                        XElement xmlabs3 = XElement.Parse(doc3.OuterXml);

                        foreach (XElement abs in xmlabs3.Elements("Item"))
                        {
                            string CounselTypeKind_for_basic = "";
                            CounselTypeKind_for_basic += abs.Attribute("name").Value;

                            //藉由查表，把輔導案件 以類型、 科別分類
                            if (counselCase_Dict.ContainsKey(CounselTypeKind_for_basic))
                            {
                                //人次
                                //沒有的話，就新增。
                                if (!counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleCountDict.ContainsKey(stuIdToDeptDict["" + interviewRecord.StudentID]))
                                {
                                    counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleCountDict.Add(stuIdToDeptDict["" + interviewRecord.StudentID], 1);

                                }
                                else
                                {   //若有的話 就加一
                                    counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleCountDict[stuIdToDeptDict["" + interviewRecord.StudentID]]++;

                                }

                                //人數
                                //沒有的話，就新增。
                                if (!counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict.ContainsKey(stuIdToDeptDict["" + interviewRecord.StudentID]))
                                {
                                    counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict.Add(stuIdToDeptDict["" + interviewRecord.StudentID], new List<string>());

                                    if (!counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict[stuIdToDeptDict["" + interviewRecord.StudentID]].Contains("" + interviewRecord.StudentID))
                                    {
                                        counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict[stuIdToDeptDict["" + interviewRecord.StudentID]].Add("" + interviewRecord.StudentID);
                                    }
                                }
                                else
                                {
                                    if (!counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict[stuIdToDeptDict["" + interviewRecord.StudentID]].Contains("" + interviewRecord.StudentID))
                                    {
                                        counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict[stuIdToDeptDict["" + interviewRecord.StudentID]].Add("" + interviewRecord.StudentID);
                                    }
                                }
                            }
                        }
                        #endregion

                        #region 老師身分類別
                        //藉由查表，把輔導案件 以類型、 科別分類
                        if (counselCase_Dict.ContainsKey(interviewRecord.authorRole))
                        {
                            //人次
                            //沒有的話，就新增。
                            if (!counselCase_Dict[interviewRecord.authorRole].CaseStaticsPeopleCountDict.ContainsKey(stuIdToDeptDict["" + interviewRecord.StudentID]))
                            {
                                counselCase_Dict[interviewRecord.authorRole].CaseStaticsPeopleCountDict.Add(stuIdToDeptDict["" + interviewRecord.StudentID], 1);

                            }
                            else
                            {   //若有的話 就加一
                                counselCase_Dict[interviewRecord.authorRole].CaseStaticsPeopleCountDict[stuIdToDeptDict["" + interviewRecord.StudentID]]++;

                            }

                            //人數
                            //沒有的話，就新增。
                            if (!counselCase_Dict[interviewRecord.authorRole].CaseStaticsPeopleDict.ContainsKey(stuIdToDeptDict["" + interviewRecord.StudentID]))
                            {
                                counselCase_Dict[interviewRecord.authorRole].CaseStaticsPeopleDict.Add(stuIdToDeptDict["" + interviewRecord.StudentID], new List<string>());

                                if (!counselCase_Dict[interviewRecord.authorRole].CaseStaticsPeopleDict[stuIdToDeptDict["" + interviewRecord.StudentID]].Contains("" + interviewRecord.StudentID))
                                {
                                    counselCase_Dict[interviewRecord.authorRole].CaseStaticsPeopleDict[stuIdToDeptDict["" + interviewRecord.StudentID]].Add("" + interviewRecord.StudentID);
                                }
                            }
                            else
                            {
                                if (!counselCase_Dict[interviewRecord.authorRole].CaseStaticsPeopleDict[stuIdToDeptDict["" + interviewRecord.StudentID]].Contains("" + interviewRecord.StudentID))
                                {
                                    counselCase_Dict[interviewRecord.authorRole].CaseStaticsPeopleDict[stuIdToDeptDict["" + interviewRecord.StudentID]].Add("" + interviewRecord.StudentID);
                                }
                            }
                        }
                        #endregion

                    }
                    #endregion


                    

                    #region 輔導教師統計
                    // 輔導案件類別
                    foreach (DAO.UDT_CounselStudentInterviewRecordDef interviewRecord in record.Value)
                    {
                        #region 輔導老師身分類別
                        //藉由查表，把有該老師ID 的 紀錄 整理
                        if (counselTeacherInterview_Dict.ContainsKey("" + "" + interviewRecord.TeacherID))
                        {
                            //人次
                            //沒有的話，就新增。
                            if (!counselTeacherInterview_Dict["" + interviewRecord.TeacherID].WaysStaticsPeopleCountDict.ContainsKey(interviewRecord.InterviewType))
                            {
                                counselTeacherInterview_Dict["" + interviewRecord.TeacherID].WaysStaticsPeopleCountDict.Add(interviewRecord.InterviewType, 1);
                            }
                            else
                            {   //若有的話 就加一
                                counselTeacherInterview_Dict["" + interviewRecord.TeacherID].WaysStaticsPeopleCountDict[interviewRecord.InterviewType]++;
                            }

                            //人數
                            //沒有的話，就新增。
                            if (!counselTeacherInterview_Dict["" + interviewRecord.TeacherID].WaysStaticsPeopleDict.ContainsKey(interviewRecord.InterviewType))
                            {
                                counselTeacherInterview_Dict["" + interviewRecord.TeacherID].WaysStaticsPeopleDict.Add(interviewRecord.InterviewType, new List<string>());

                                if (!counselTeacherInterview_Dict["" + interviewRecord.TeacherID].WaysStaticsPeopleDict[interviewRecord.InterviewType].Contains("" + interviewRecord.StudentID))
                                {
                                    counselTeacherInterview_Dict["" + interviewRecord.TeacherID].WaysStaticsPeopleDict[interviewRecord.InterviewType].Add("" + interviewRecord.StudentID);
                                }
                            }
                            else
                            {
                                if (!counselTeacherInterview_Dict["" + interviewRecord.TeacherID].WaysStaticsPeopleDict[interviewRecord.InterviewType].Contains("" + interviewRecord.StudentID))
                                {
                                    counselTeacherInterview_Dict["" + interviewRecord.TeacherID].WaysStaticsPeopleDict[interviewRecord.InterviewType].Add("" + interviewRecord.StudentID);
                                }
                            }
                        }
                        #endregion

                    }
                    #endregion

                    #region 老師(高一、高二、高三)統計 

                    // 輔導案件類別
                    foreach (DAO.UDT_CounselStudentInterviewRecordDef interviewRecord in record.Value)
                    {
                        #region 輔導老師身分類別
                        // 只統計班導
                        if (interviewRecord.authorRole == "班導師")
                        {
                            //藉由查表，把有該老師ID 的 紀錄 整理
                            if (TeacherInterview_Dict.ContainsKey(stuIdToGradeDict["" + interviewRecord.StudentID]))
                            {
                                //人次
                                //沒有的話，就新增。
                                if (!TeacherInterview_Dict[stuIdToGradeDict["" + interviewRecord.StudentID]].WaysStaticsPeopleCountDict.ContainsKey(interviewRecord.InterviewType))
                                {
                                    TeacherInterview_Dict[stuIdToGradeDict["" + interviewRecord.StudentID]].WaysStaticsPeopleCountDict.Add(interviewRecord.InterviewType, 1);
                                }
                                else
                                {   //若有的話 就加一
                                    TeacherInterview_Dict[stuIdToGradeDict["" + interviewRecord.StudentID]].WaysStaticsPeopleCountDict[interviewRecord.InterviewType]++;
                                }

                                //人數
                                //沒有的話，就新增。
                                if (!TeacherInterview_Dict[stuIdToGradeDict["" + interviewRecord.StudentID]].WaysStaticsPeopleDict.ContainsKey(interviewRecord.InterviewType))
                                {
                                    TeacherInterview_Dict[stuIdToGradeDict["" + interviewRecord.StudentID]].WaysStaticsPeopleDict.Add(interviewRecord.InterviewType, new List<string>());

                                    if (!TeacherInterview_Dict[stuIdToGradeDict["" + interviewRecord.StudentID]].WaysStaticsPeopleDict[interviewRecord.InterviewType].Contains("" + interviewRecord.StudentID))
                                    {
                                        TeacherInterview_Dict[stuIdToGradeDict["" + interviewRecord.StudentID]].WaysStaticsPeopleDict[interviewRecord.InterviewType].Add("" + interviewRecord.StudentID);
                                    }
                                }
                                else
                                {
                                    if (!TeacherInterview_Dict[stuIdToGradeDict["" + interviewRecord.StudentID]].WaysStaticsPeopleDict[interviewRecord.InterviewType].Contains("" + interviewRecord.StudentID))
                                    {
                                        TeacherInterview_Dict[stuIdToGradeDict["" + interviewRecord.StudentID]].WaysStaticsPeopleDict[interviewRecord.InterviewType].Add("" + interviewRecord.StudentID);
                                    }
                                }
                            }

                        }


                        #endregion

                    }
                    #endregion

                }


                //聯繫紀錄
                foreach (KeyValuePair<string, List<DAO.UDT_Counsel_home_visit_RecordDef>> record in dicHomeVisitRecord)
                {
                    #region 導師統計(年級)
                    //一年級老師
                    foreach (DAO.AllCounselStaticticsTeacherRecord teacherRecord in CounselTeacherList_grade1)
                    {
                        if (record.Key == teacherRecord.TeacherID)
                        {
                            //人數
                            teacherRecord.CounselPeople = GetCounselPeople(record.Key, record.Value);
                            //加上人次
                            teacherRecord.CounselPeopleCount += record.Value.Count();
                        }
                    }
                    //二年級老師
                    foreach (DAO.AllCounselStaticticsTeacherRecord teacherRecord in CounselTeacherList_grade2)
                    {
                        if (record.Key == teacherRecord.TeacherID)
                        {
                            //人數
                            teacherRecord.CounselPeople = GetCounselPeople(record.Key, record.Value);
                            //加上人次
                            teacherRecord.CounselPeopleCount += record.Value.Count();
                        }
                    }
                    //三年級老師
                    foreach (DAO.AllCounselStaticticsTeacherRecord teacherRecord in CounselTeacherList_grade3)
                    {
                        if (record.Key == teacherRecord.TeacherID)
                        {
                            //人數
                            teacherRecord.CounselPeople = GetCounselPeople(record.Key, record.Value);
                            //加上人次
                            teacherRecord.CounselPeopleCount += record.Value.Count();
                        }
                    }
                    #endregion

                    #region 輔導案件類別
                    // 輔導案件類別
                    foreach (DAO.UDT_Counsel_home_visit_RecordDef homeVisitRecord in record.Value)
                    {

                        #region 解析輔導類型Xml
                        XmlDocument doc3 = new XmlDocument();
                        //幫忙加根目錄
                        string xmlContent3 = "<root>" + homeVisitRecord.CounselTypeKind + "</root>";
                        doc3.LoadXml(xmlContent3);
                        XmlNode newNode3 = doc3.DocumentElement;
                        doc3.AppendChild(newNode3);
                        XElement xmlabs3 = XElement.Parse(doc3.OuterXml);

                        foreach (XElement abs in xmlabs3.Elements("Item"))
                        {
                            string CounselTypeKind_for_basic = "";
                            CounselTypeKind_for_basic += abs.Attribute("name").Value;

                            //藉由查表，把輔導案件 以類型、 科別分類
                            if (counselCase_Dict.ContainsKey(CounselTypeKind_for_basic))
                            {
                                //人次
                                //沒有的話，就新增。
                                if (!counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleCountDict.ContainsKey(stuIdToDeptDict["" + homeVisitRecord.StudentID]))
                                {
                                    counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleCountDict.Add(stuIdToDeptDict["" + homeVisitRecord.StudentID], 1);

                                }
                                else
                                {   //若有的話 就加一
                                    counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleCountDict[stuIdToDeptDict["" + homeVisitRecord.StudentID]]++;

                                }

                                //人數
                                //沒有的話，就新增。
                                if (!counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict.ContainsKey(stuIdToDeptDict["" + homeVisitRecord.StudentID]))
                                {
                                    counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict.Add(stuIdToDeptDict["" + homeVisitRecord.StudentID], new List<string>());

                                    if (!counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict[stuIdToDeptDict["" + homeVisitRecord.StudentID]].Contains("" + homeVisitRecord.StudentID))
                                    {
                                        counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict[stuIdToDeptDict["" + homeVisitRecord.StudentID]].Add("" + homeVisitRecord.StudentID);
                                    }
                                }
                                else
                                {
                                    if (!counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict[stuIdToDeptDict["" + homeVisitRecord.StudentID]].Contains("" + homeVisitRecord.StudentID))
                                    {
                                        counselCase_Dict[CounselTypeKind_for_basic].CaseStaticsPeopleDict[stuIdToDeptDict["" + homeVisitRecord.StudentID]].Add("" + homeVisitRecord.StudentID);
                                    }
                                }
                            }
                        }
                        #endregion

                        #region 老師身分類別
                        //藉由查表，把輔導案件 以類型、 科別分類
                        if (counselCase_Dict.ContainsKey(homeVisitRecord.authorRole))
                        {
                            //人次
                            //沒有的話，就新增。
                            if (!counselCase_Dict[homeVisitRecord.authorRole].CaseStaticsPeopleCountDict.ContainsKey(stuIdToDeptDict["" + homeVisitRecord.StudentID]))
                            {
                                counselCase_Dict[homeVisitRecord.authorRole].CaseStaticsPeopleCountDict.Add(stuIdToDeptDict["" + homeVisitRecord.StudentID], 1);

                            }
                            else
                            {   //若有的話 就加一
                                counselCase_Dict[homeVisitRecord.authorRole].CaseStaticsPeopleCountDict[stuIdToDeptDict["" + homeVisitRecord.StudentID]]++;

                            }

                            //人數
                            //沒有的話，就新增。
                            if (!counselCase_Dict[homeVisitRecord.authorRole].CaseStaticsPeopleDict.ContainsKey(stuIdToDeptDict["" + homeVisitRecord.StudentID]))
                            {
                                counselCase_Dict[homeVisitRecord.authorRole].CaseStaticsPeopleDict.Add(stuIdToDeptDict["" + homeVisitRecord.StudentID], new List<string>());

                                if (!counselCase_Dict[homeVisitRecord.authorRole].CaseStaticsPeopleDict[stuIdToDeptDict["" + homeVisitRecord.StudentID]].Contains("" + homeVisitRecord.StudentID))
                                {
                                    counselCase_Dict[homeVisitRecord.authorRole].CaseStaticsPeopleDict[stuIdToDeptDict["" + homeVisitRecord.StudentID]].Add("" + homeVisitRecord.StudentID);
                                }
                            }
                            else
                            {
                                if (!counselCase_Dict[homeVisitRecord.authorRole].CaseStaticsPeopleDict[stuIdToDeptDict["" + homeVisitRecord.StudentID]].Contains("" + homeVisitRecord.StudentID))
                                {
                                    counselCase_Dict[homeVisitRecord.authorRole].CaseStaticsPeopleDict[stuIdToDeptDict["" + homeVisitRecord.StudentID]].Add("" + homeVisitRecord.StudentID);
                                }
                            }
                        }
                        #endregion

                    }
                    #endregion

                    #region 輔導教師統計
                    // 輔導案件類別
                    foreach (DAO.UDT_Counsel_home_visit_RecordDef homeVisitRecord in record.Value)
                    {
                        #region 輔導老師身分類別
                        //藉由查表，把有該老師ID 的 紀錄 整理
                        if (counselTeacherHomeVisit_Dict.ContainsKey("" + "" + homeVisitRecord.TeacherID))
                        {
                            //人次
                            //沒有的話，就新增。
                            if (!counselTeacherHomeVisit_Dict["" + homeVisitRecord.TeacherID].WaysStaticsPeopleCountDict.ContainsKey(homeVisitRecord.home_visit_type))
                            {
                                counselTeacherHomeVisit_Dict["" + homeVisitRecord.TeacherID].WaysStaticsPeopleCountDict.Add(homeVisitRecord.home_visit_type, 1);
                            }
                            else
                            {   //若有的話 就加一
                                counselTeacherHomeVisit_Dict["" + homeVisitRecord.TeacherID].WaysStaticsPeopleCountDict[homeVisitRecord.home_visit_type]++;
                            }

                            //人數
                            //沒有的話，就新增。
                            if (!counselTeacherHomeVisit_Dict["" + homeVisitRecord.TeacherID].WaysStaticsPeopleDict.ContainsKey(homeVisitRecord.home_visit_type))
                            {
                                counselTeacherHomeVisit_Dict["" + homeVisitRecord.TeacherID].WaysStaticsPeopleDict.Add(homeVisitRecord.home_visit_type, new List<string>());

                                if (!counselTeacherHomeVisit_Dict["" + homeVisitRecord.TeacherID].WaysStaticsPeopleDict[homeVisitRecord.home_visit_type].Contains("" + homeVisitRecord.StudentID))
                                {
                                    counselTeacherHomeVisit_Dict["" + homeVisitRecord.TeacherID].WaysStaticsPeopleDict[homeVisitRecord.home_visit_type].Add("" + homeVisitRecord.StudentID);
                                }
                            }
                            else
                            {
                                if (!counselTeacherHomeVisit_Dict["" + homeVisitRecord.TeacherID].WaysStaticsPeopleDict[homeVisitRecord.home_visit_type].Contains("" + homeVisitRecord.StudentID))
                                {
                                    counselTeacherHomeVisit_Dict["" + homeVisitRecord.TeacherID].WaysStaticsPeopleDict[homeVisitRecord.home_visit_type].Add("" + homeVisitRecord.StudentID);
                                }
                            }
                        }
                        #endregion

                    }
                    #endregion

                    #region 老師(高一、高二、高三)統計 
                    
                    // 輔導案件類別
                    foreach (DAO.UDT_Counsel_home_visit_RecordDef homeVisitRecord in record.Value)
                    {
                        #region 輔導老師身分類別
                        // 只統計班導
                        if (homeVisitRecord.authorRole == "班導師")
                        {
                            //藉由查表，把有該老師ID 的 紀錄 整理
                            if (TeacherHomeVisit_Dict.ContainsKey(stuIdToGradeDict["" + homeVisitRecord.StudentID]))
                            {
                                //人次
                                //沒有的話，就新增。
                                if (!TeacherHomeVisit_Dict[stuIdToGradeDict["" + homeVisitRecord.StudentID]].WaysStaticsPeopleCountDict.ContainsKey(homeVisitRecord.home_visit_type))
                                {
                                    TeacherHomeVisit_Dict[stuIdToGradeDict["" + homeVisitRecord.StudentID]].WaysStaticsPeopleCountDict.Add(homeVisitRecord.home_visit_type, 1);
                                }
                                else
                                {   //若有的話 就加一
                                    TeacherHomeVisit_Dict[stuIdToGradeDict["" + homeVisitRecord.StudentID]].WaysStaticsPeopleCountDict[homeVisitRecord.home_visit_type]++;
                                }

                                //人數
                                //沒有的話，就新增。
                                if (!TeacherHomeVisit_Dict[stuIdToGradeDict["" + homeVisitRecord.StudentID]].WaysStaticsPeopleDict.ContainsKey(homeVisitRecord.home_visit_type))
                                {
                                    TeacherHomeVisit_Dict[stuIdToGradeDict["" + homeVisitRecord.StudentID]].WaysStaticsPeopleDict.Add(homeVisitRecord.home_visit_type, new List<string>());

                                    if (!TeacherHomeVisit_Dict[stuIdToGradeDict["" + homeVisitRecord.StudentID]].WaysStaticsPeopleDict[homeVisitRecord.home_visit_type].Contains("" + homeVisitRecord.StudentID))
                                    {
                                        TeacherHomeVisit_Dict[stuIdToGradeDict["" + homeVisitRecord.StudentID]].WaysStaticsPeopleDict[homeVisitRecord.home_visit_type].Add("" + homeVisitRecord.StudentID);
                                    }
                                }
                                else
                                {
                                    if (!TeacherHomeVisit_Dict[stuIdToGradeDict["" + homeVisitRecord.StudentID]].WaysStaticsPeopleDict[homeVisitRecord.home_visit_type].Contains("" + homeVisitRecord.StudentID))
                                    {
                                        TeacherHomeVisit_Dict[stuIdToGradeDict["" + homeVisitRecord.StudentID]].WaysStaticsPeopleDict[homeVisitRecord.home_visit_type].Add("" + homeVisitRecord.StudentID);
                                    }
                                }
                            }

                        }

                        
                        #endregion

                    }
                    #endregion
                }





                int row_index = 1;

                DataTable dt = new DataTable();
                var basicDataTitleList = new List<string>();
                var attendeesDataTitleList = new List<string>();
                var counselTypeDataTitleList = new List<string>();
                var counselTypeKindDataTitleList = new List<string>();
                #region 基本欄位
                //所有Datatable 欄位--單項目
                dt.Columns.Add("學號");
                dt.Columns.Add("學生姓名");
                dt.Columns.Add("性別");
                dt.Columns.Add("班級");
                dt.Columns.Add("座號");
                dt.Columns.Add("晤談編號");
                dt.Columns.Add("學年度");
                dt.Columns.Add("學期");
                dt.Columns.Add("晤談日期");
                dt.Columns.Add("晤談時間");
                dt.Columns.Add("晤談動機");
                dt.Columns.Add("輔導對象");
                dt.Columns.Add("晤談方式");
                dt.Columns.Add("地點");
                dt.Columns.Add("內容要點");
                dt.Columns.Add("記錄人的登入帳號");
                dt.Columns.Add("記錄人的姓名");
                dt.Columns.Add("是否公開");
                dt.Columns.Add("記錄人");

                //所有Datatable 欄位--多項目
                dt.Columns.Add("參與人員");
                dt.Columns.Add("處理方式");
                dt.Columns.Add("案件類別");

                //基本資料Title--單項目
                basicDataTitleList.Add("學號");
                basicDataTitleList.Add("學生姓名");
                basicDataTitleList.Add("性別");
                basicDataTitleList.Add("班級");
                basicDataTitleList.Add("座號");
                basicDataTitleList.Add("晤談編號");
                basicDataTitleList.Add("學年度");
                basicDataTitleList.Add("學期");
                basicDataTitleList.Add("晤談日期");
                basicDataTitleList.Add("晤談時間");
                basicDataTitleList.Add("晤談動機");
                basicDataTitleList.Add("輔導對象");
                basicDataTitleList.Add("晤談方式");
                basicDataTitleList.Add("地點");

                //基本資料Title--多項目
                basicDataTitleList.Add("參與人員");
                basicDataTitleList.Add("處理方式");
                basicDataTitleList.Add("案件類別");

                basicDataTitleList.Add("內容要點");
                //basicDataTitleList.Add("記錄人的登入帳號");
                basicDataTitleList.Add("記錄人的姓名");
                basicDataTitleList.Add("是否公開");
                basicDataTitleList.Add("記錄人");
                #endregion

                progressCount = 0;
                //開始整理資料


                int col_index = 3;

                #region 加入Excel 表單 每欄Title

                wb.Open(new MemoryStream(Properties.Resources.輔導案量統計空白樣板), FileFormatType.Excel2003);

                Cells cs_classTeacher_template = wb.Worksheets["導師統計(樣板)"].Cells;

                Cells cs_CounselTypeKind_template = wb.Worksheets["案件類別(樣板)"].Cells;


                Cells cs_classTeacher1 = wb.Worksheets["導師統計(一年級)"].Cells;
                Cells cs_classTeacher2 = wb.Worksheets["導師統計(二年級)"].Cells;
                Cells cs_classTeacher3 = wb.Worksheets["導師統計(三年級)"].Cells;

                Cells cs_CounselTypeKind = wb.Worksheets["案件類別"].Cells;

                Cells cs_CounselTeacherWays = wb.Worksheets["輔導教師統計"].Cells;

                Cells cs_TeacherWays = wb.Worksheets["導師統計"].Cells;


                // 科別表格
                Range deparment = cs_classTeacher_template.CreateRange(3, 2, false);

                //合計表格
                Range totalStatistics = cs_classTeacher_template.CreateRange(5, 2, false);

                // 科別表格(案件類別)
                Range deparment_CounselTypeKind = cs_CounselTypeKind_template.CreateRange(3, 2, false);

                //合計表格(案件類別)
                Range totalStatistics_CounselTypeKind = cs_CounselTypeKind_template.CreateRange(5, 2, false);

                //統計 各年級有幾個科別
                List<string> departmentList1 = new List<string>();
                List<string> departmentList2 = new List<string>();
                List<string> departmentList3 = new List<string>();

                //建立各科別對應的行數位置
                Dictionary<string, int> departmentRow1 = new Dictionary<string, int>();
                Dictionary<string, int> departmentRow2 = new Dictionary<string, int>();
                Dictionary<string, int> departmentRow3 = new Dictionary<string, int>();

                //統計 全年級有幾個科別
                List<string> departmentList_CounselTypeKind = new List<string>();

                //建立各科別對應的行數位置(案件類別)
                Dictionary<string, int> departmentRow1_CounselTypeKind = new Dictionary<string, int>();

                #region 導師統計

                #region 一年級
                foreach (DAO.AllCounselStaticticsTeacherRecord acstr in CounselTeacherList_grade1)
                {
                    if (!departmentList1.Contains(acstr.ClassDepartment))
                    {
                        departmentList1.Add(acstr.ClassDepartment);
                    }
                }

                for (int i = 0; i < departmentList1.Count; i++)
                {
                    //新增科別表格區域
                    cs_classTeacher1.CreateRange(3 + i * 2, 2, false).Copy(deparment);
                    //填上科別
                    cs_classTeacher1[3 + i * 2, 0].PutValue(departmentList1[i]);
                    //建立科別對應行數
                    departmentRow1.Add(departmentList1[i], 3 + i * 2);
                }

                //新增 合計表格區域
                cs_classTeacher1.CreateRange(3 + departmentList1.Count * 2, 2, false).Copy(totalStatistics);

                //補上合計公式
                for (int i = 0; i < CounselTeacherList_grade1.Count; i++)
                {
                    cs_classTeacher1[3 + departmentList1.Count * 2, i + 3].Formula = GetCounselPeopleTotalSumFormula(departmentList1, i + 3);
                    cs_classTeacher1[4 + departmentList1.Count * 2, i + 3].Formula = GetCounselPeopleCountTotalSumFormula(departmentList1, i + 3);
                }

                #endregion

                #region 二年級
                foreach (DAO.AllCounselStaticticsTeacherRecord acstr in CounselTeacherList_grade2)
                {
                    if (!departmentList2.Contains(acstr.ClassDepartment))
                    {
                        departmentList2.Add(acstr.ClassDepartment);
                    }
                }

                for (int i = 0; i < departmentList2.Count; i++)
                {
                    //新增科別表格區域
                    cs_classTeacher2.CreateRange(3 + i * 2, 2, false).Copy(deparment);
                    //填上科別
                    cs_classTeacher2[3 + i * 2, 0].PutValue(departmentList2[i]);
                    //建立科別對應行數
                    departmentRow2.Add(departmentList2[i], 3 + i * 2);
                }

                //新增 合計表格區域
                cs_classTeacher2.CreateRange(3 + departmentList2.Count * 2, 2, false).Copy(totalStatistics);

                //補上合計公式
                for (int i = 0; i < CounselTeacherList_grade2.Count; i++)
                {
                    cs_classTeacher2[3 + departmentList2.Count * 2, i + 3].Formula = GetCounselPeopleTotalSumFormula(departmentList2, i + 3);
                    cs_classTeacher2[4 + departmentList2.Count * 2, i + 3].Formula = GetCounselPeopleCountTotalSumFormula(departmentList2, i + 3);
                }
                #endregion

                #region 三年級
                foreach (DAO.AllCounselStaticticsTeacherRecord acstr in CounselTeacherList_grade3)
                {
                    if (!departmentList3.Contains(acstr.ClassDepartment))
                    {
                        departmentList3.Add(acstr.ClassDepartment);
                    }
                }

                for (int i = 0; i < departmentList3.Count; i++)
                {
                    //新增科別表格區域
                    cs_classTeacher3.CreateRange(3 + i * 2, 2, false).Copy(deparment);
                    //填上科別
                    cs_classTeacher3[3 + i * 2, 0].PutValue(departmentList3[i]);
                    //建立科別對應行數
                    departmentRow3.Add(departmentList3[i], 3 + i * 2);
                }

                //新增 合計表格區域
                cs_classTeacher3.CreateRange(3 + departmentList3.Count * 2, 2, false).Copy(totalStatistics);

                //補上合計公式
                for (int i = 0; i < CounselTeacherList_grade3.Count; i++)
                {
                    cs_classTeacher3[3 + departmentList3.Count * 2, i + 3].Formula = GetCounselPeopleTotalSumFormula(departmentList3, i + 3);
                    cs_classTeacher3[4 + departmentList3.Count * 2, i + 3].Formula = GetCounselPeopleCountTotalSumFormula(departmentList3, i + 3);
                }

                #endregion


                #region 一年級填值
                foreach (DAO.AllCounselStaticticsTeacherRecord acstr in CounselTeacherList_grade1)
                {
                    // 表頭
                    cs_classTeacher1[0, 0].PutValue("臺中市立文華高級中等學校" + defaultYear + "學年度第" + defaultSemester + "學期 學生家庭聯繫＆個人晤談情況統計表 \r\n（" + filterDateBegin.ToShortDateString() + "～" + filterDateEnd.ToShortDateString() + "）");
                    //班級名稱
                    cs_classTeacher1[1, col_index].PutValue(acstr.ClassName);
                    //教師名稱
                    cs_classTeacher1[2, col_index].PutValue(acstr.TeacherName);
                    //人數
                    cs_classTeacher1[departmentRow1[acstr.ClassDepartment], col_index].PutValue(acstr.CounselPeople);
                    //人次
                    cs_classTeacher1[departmentRow1[acstr.ClassDepartment] + 1, col_index].PutValue(acstr.CounselPeopleCount);

                    col_index++;
                }
                #endregion

                #region 二年級填值
                // 將col_index 歸正
                col_index = 3;

                foreach (DAO.AllCounselStaticticsTeacherRecord acstr in CounselTeacherList_grade2)
                {
                    // 表頭
                    cs_classTeacher2[0, 0].PutValue("臺中市立文華高級中等學校" + defaultYear + "學年度第" + defaultSemester + "學期 學生家庭聯繫＆個人晤談情況統計表 \r\n（" + filterDateBegin.ToShortDateString() + "～" + filterDateEnd.ToShortDateString() + "）");
                    //班級名稱
                    cs_classTeacher2[1, col_index].PutValue(acstr.ClassName);
                    //教師名稱
                    cs_classTeacher2[2, col_index].PutValue(acstr.TeacherName);
                    //人數
                    cs_classTeacher2[departmentRow2[acstr.ClassDepartment], col_index].PutValue(acstr.CounselPeople);
                    //人次
                    cs_classTeacher2[departmentRow2[acstr.ClassDepartment] + 1, col_index].PutValue(acstr.CounselPeopleCount);

                    col_index++;
                }

                #endregion

                #region 三年級填值
                // 將col_index 歸正
                col_index = 3;

                foreach (DAO.AllCounselStaticticsTeacherRecord acstr in CounselTeacherList_grade3)
                {
                    // 表頭
                    cs_classTeacher3[0, 0].PutValue("臺中市立文華高級中等學校" + defaultYear + "學年度第" + defaultSemester + "學期 學生家庭聯繫＆個人晤談情況統計表 \r\n（" + filterDateBegin.ToShortDateString() + "～" + filterDateEnd.ToShortDateString() + "）");
                    //班級名稱
                    cs_classTeacher3[1, col_index].PutValue(acstr.ClassName);
                    //教師名稱
                    cs_classTeacher3[2, col_index].PutValue(acstr.TeacherName);
                    //人數
                    cs_classTeacher3[departmentRow3[acstr.ClassDepartment], col_index].PutValue(acstr.CounselPeople);
                    //人次
                    cs_classTeacher3[departmentRow3[acstr.ClassDepartment] + 1, col_index].PutValue(acstr.CounselPeopleCount);

                    col_index++;
                }
                #endregion

                #endregion

                #region 案件類別

                //將三個年級 科別 統一統計
                #region 彙整科別
                foreach (string dept in departmentList1)
                {
                    if (!departmentList_CounselTypeKind.Contains(dept))
                    {
                        departmentList_CounselTypeKind.Add(dept);
                    }
                }
                foreach (string dept in departmentList2)
                {
                    if (!departmentList_CounselTypeKind.Contains(dept))
                    {
                        departmentList_CounselTypeKind.Add(dept);
                    }
                }
                foreach (string dept in departmentList3)
                {
                    if (!departmentList_CounselTypeKind.Contains(dept))
                    {
                        departmentList_CounselTypeKind.Add(dept);
                    }
                }
                #endregion

                #region 表格建立
                for (int i = 0; i < departmentList_CounselTypeKind.Count; i++)
                {
                    //新增科別表格區域
                    cs_CounselTypeKind.CreateRange(3 + i * 2, 2, false).Copy(deparment_CounselTypeKind);
                    //填上科別
                    cs_CounselTypeKind[3 + i * 2, 0].PutValue(departmentList_CounselTypeKind[i]);
                    //建立科別對應行數
                    departmentRow1_CounselTypeKind.Add(departmentList_CounselTypeKind[i], 3 + i * 2);
                }

                //新增 合計表格區域
                cs_CounselTypeKind.CreateRange(3 + departmentList_CounselTypeKind.Count * 2, 2, false).Copy(totalStatistics_CounselTypeKind);

                //補上合計公式
                for (int i = 0; i < counselCase_Dict.Count; i++)
                {
                    cs_CounselTypeKind[3 + departmentList_CounselTypeKind.Count * 2, i + 3].Formula = GetCounselPeopleTotalSumFormula(departmentList_CounselTypeKind, i + 3);
                    cs_CounselTypeKind[4 + departmentList_CounselTypeKind.Count * 2, i + 3].Formula = GetCounselPeopleCountTotalSumFormula(departmentList_CounselTypeKind, i + 3);
                }
                #endregion

                // 將col_index 歸正
                col_index = 3;
                //填值
                foreach (KeyValuePair<string, DAO.AllCounselCaseStaticticsRecord> record in counselCase_Dict)
                {
                    // 表頭
                    cs_CounselTypeKind[0, 0].PutValue("臺中市立文華高級中等學校" + defaultYear + "學年度第" + defaultSemester + "學期 學生家庭聯繫＆個人晤談情況統計表 \r\n（" + filterDateBegin.ToShortDateString() + "～" + filterDateEnd.ToShortDateString() + "）");

                    
                    //人數
                    foreach (KeyValuePair<string,List<string>> r1 in record.Value.CaseStaticsPeopleDict)
                    {
                        cs_CounselTypeKind[departmentRow1_CounselTypeKind[r1.Key], col_index].PutValue(r1.Value.Count());
                    }
                    //人次
                    foreach (KeyValuePair<string, int> r1 in record.Value.CaseStaticsPeopleCountDict)
                    {
                        cs_CounselTypeKind[departmentRow1_CounselTypeKind[r1.Key]+1, col_index].PutValue(r1.Value);
                    }
                   
                    col_index++;
                }


                #endregion



                #region 輔導教師統計
                //輔導教師統計
                
                //家訪紀錄
                row_index = 3; // 從第4列 開始填

                foreach (KeyValuePair<string, DAO.CounselTeacherHomeVisitWaysRecord> record in counselTeacherHomeVisit_Dict)
                {
                    col_index = 1;  // 從第2行 開始填

                    //表頭
                    cs_CounselTeacherWays[0, 0].PutValue("家庭教育自我檢核統計表-輔導處   \r\n 統計時間:" + filterDateBegin.ToShortDateString() + "～" + filterDateEnd.ToShortDateString());

                    //輔導老師姓名
                    cs_CounselTeacherWays[row_index, 0].PutValue(record.Value.CounselTeacherName);

                    //填上人數
                    foreach (KeyValuePair<string, List<string>> r1 in record.Value.WaysStaticsPeopleDict)
                    {
                        cs_CounselTeacherWays[row_index, col_index].PutValue(r1.Value.Count());

                        col_index = col_index + 2;
                    }

                    col_index = 2; // 從第3行 開始填
                    //填上人次
                    foreach (KeyValuePair<string, int> r1 in record.Value.WaysStaticsPeopleCountDict)
                    {
                        cs_CounselTeacherWays[row_index, col_index].PutValue(r1.Value);

                        col_index = col_index + 2;
                    }

                    row_index++;
                }




                // 晤談紀錄
                row_index = 15; // 從第16列 開始填

                foreach (KeyValuePair<string, DAO.CounselTeacherInterviewWaysRecord> record in counselTeacherInterview_Dict)
                {
                    col_index = 1;  // 從第2行 開始填

                    //表頭
                    cs_CounselTeacherWays[12, 0].PutValue("學生晤談自我檢核統計表-輔導處  \r\n 統計時間:" + filterDateBegin.ToShortDateString() + "～" + filterDateEnd.ToShortDateString());

                    //輔導老師姓名
                    cs_CounselTeacherWays[row_index, 0].PutValue(record.Value.CounselTeacherName);

                    //填上人數
                    foreach (KeyValuePair<string, List<string>> r1 in record.Value.WaysStaticsPeopleDict)
                    {
                        cs_CounselTeacherWays[row_index, col_index].PutValue(r1.Value.Count());

                        col_index = col_index + 2;
                    }


                    col_index = 2; // 從第3行 開始填
                    //填上人次
                    foreach (KeyValuePair<string, int> r1 in record.Value.WaysStaticsPeopleCountDict)
                    {
                        cs_CounselTeacherWays[row_index, col_index].PutValue(r1.Value);

                        col_index = col_index + 2;
                    }

                    row_index++;
                }
                #endregion

                

                #region 老師(高一、高二、高三)統計
                //輔導教師統計

                //家訪紀錄
                row_index = 3; // 從第4列 開始填

                foreach (KeyValuePair<string, DAO.TeacherHomeVisitWaysRecord> record in TeacherHomeVisit_Dict)
                {
                    col_index = 1;  // 從第2行 開始填

                    //表頭
                    cs_TeacherWays[0, 0].PutValue("家庭教育自我檢核統計表-導師   \r\n 統計時間:" + filterDateBegin.ToShortDateString() + "～" + filterDateEnd.ToShortDateString());

                    //輔導老師姓名
                    cs_TeacherWays[row_index, 0].PutValue(record.Value.CounselTeacherName);

                    //填上人數
                    foreach (KeyValuePair<string, List<string>> r1 in record.Value.WaysStaticsPeopleDict)
                    {
                        cs_TeacherWays[row_index, col_index].PutValue(r1.Value.Count());

                        col_index = col_index + 2;
                    }

                    col_index = 2; // 從第3行 開始填
                    //填上人次
                    foreach (KeyValuePair<string, int> r1 in record.Value.WaysStaticsPeopleCountDict)
                    {
                        cs_TeacherWays[row_index, col_index].PutValue(r1.Value);

                        col_index = col_index + 2;
                    }

                    row_index++;
                }




                // 晤談紀錄
                row_index = 12; // 從第13列 開始填

                foreach (KeyValuePair<string, DAO.CounselTeacherInterviewRecord> record in TeacherInterview_Dict)
                {
                    col_index = 1;  // 從第2行 開始填

                    //表頭
                    cs_TeacherWays[9, 0].PutValue("學生晤談自我檢核統計表-導師  \r\n 統計時間:" + filterDateBegin.ToShortDateString() + "～" + filterDateEnd.ToShortDateString());

                    //輔導老師姓名
                    cs_TeacherWays[row_index, 0].PutValue(record.Value.CounselTeacherName);

                    //填上人數
                    foreach (KeyValuePair<string, List<string>> r1 in record.Value.WaysStaticsPeopleDict)
                    {
                        cs_TeacherWays[row_index, col_index].PutValue(r1.Value.Count());

                        col_index = col_index + 2;
                    }


                    col_index = 2; // 從第3行 開始填
                    //填上人次
                    foreach (KeyValuePair<string, int> r1 in record.Value.WaysStaticsPeopleCountDict)
                    {
                        cs_TeacherWays[row_index, col_index].PutValue(r1.Value);

                        col_index = col_index + 2;
                    }

                    row_index++;
                }
                #endregion

                //移除樣板

                wb.Worksheets.RemoveAt("導師統計(樣板)");
                wb.Worksheets.RemoveAt("案件類別(樣板)");

                #endregion

             

                #region 舊列印方式
                //foreach (var student_id in student_id_List) 
                //{
                //    StudentRecord stuRec = new StudentRecord();

                //    if (InterviewRecord_Dict.ContainsKey(student_id))
                //    {
                //        if (studentRecord_Dict.ContainsKey(student_id)) 
                //        {
                //            stuRec = studentRecord_Dict[student_id];
                //        }

                //        foreach (var InterviewRecord in InterviewRecord_Dict[student_id]) 
                //        {
                //            //學生姓名
                //            cs[row_index, 0].PutValue(stuRec.Name);
                //            //學號
                //            cs[row_index, 1].PutValue(stuRec.StudentNumber);
                //            //性別
                //            cs[row_index, 2].PutValue(stuRec.Gender);
                //            //班級
                //            cs[row_index, 3].PutValue(  stuRec.Class !=null? stuRec.Class.Name:"");
                //            //座號
                //            cs[row_index, 4].PutValue(stuRec.SeatNo);
                //            //晤談編號
                //            cs[row_index, 5].PutValue(InterviewRecord.InterviewNo);
                //            //學年度
                //            cs[row_index, 6].PutValue(InterviewRecord.schoolyear);
                //            //學期
                //            cs[row_index, 7].PutValue(InterviewRecord.semester);
                //            //晤談日期
                //            cs[row_index, 8].PutValue(InterviewRecord.InterviewDate.Value.ToShortDateString());
                //            //晤談時間
                //            cs[row_index, 9].PutValue(InterviewRecord.InterviewTime);
                //            //晤談動機
                //            cs[row_index, 10].PutValue(InterviewRecord.Cause);
                //            //輔導對象
                //            cs[row_index, 11].PutValue(InterviewRecord.IntervieweeType);
                //            //晤談方式
                //            cs[row_index, 12].PutValue(InterviewRecord.InterviewType);
                //            //地點
                //            cs[row_index, 13].PutValue(InterviewRecord.Place);

                //            XmlDocument doc1 = new XmlDocument();
                //            //幫忙加根目錄
                //            string xmlContent1 = "<root>" + InterviewRecord.Attendees + "</root>";
                //            doc1.LoadXml(xmlContent1);
                //            XmlNode newNode1 = doc1.DocumentElement;
                //            doc1.AppendChild(newNode1);
                //            XElement xmlabs1 = XElement.Parse(doc1.OuterXml);
                //            string attendees = "";

                //            foreach (XElement abs in xmlabs1.Elements("Item"))
                //            {
                //              attendees += abs.Attribute("name").Value;
                //              if (abs != xmlabs1.LastNode) 
                //              {
                //                  attendees += "、";                          
                //              }                         
                //            }

                //            //參與人員
                //            cs[row_index, 14].PutValue(attendees);

                //            XmlDocument doc2 = new XmlDocument();
                //            //幫忙加根目錄
                //            string xmlContent2 = "<root>" + InterviewRecord.CounselType + "</root>";
                //            doc2.LoadXml(xmlContent2);
                //            XmlNode newNode2 = doc2.DocumentElement;
                //            doc2.AppendChild(newNode2);
                //            XElement xmlabs2 = XElement.Parse(doc2.OuterXml);
                //            string CounselType = "";

                //            foreach (XElement abs in xmlabs2.Elements("Item"))
                //            {
                //                CounselType += abs.Attribute("name").Value;
                //                if (abs.Attribute("name").Value == "其他") 
                //                {
                //                    if(abs.Attribute("remark")!=null)
                //                    CounselType += ":"+abs.Attribute("remark").Value;                            
                //                }
                //                if (abs != xmlabs2.LastNode)
                //                {
                //                    CounselType += "、";
                //                }
                //            }

                //            //處理方式
                //            cs[row_index, 15].PutValue(CounselType);
                //            XmlDocument doc3 = new XmlDocument();
                //            //幫忙加根目錄
                //            string xmlContent3 = "<root>" + InterviewRecord.CounselTypeKind + "</root>";
                //            doc3.LoadXml(xmlContent3);
                //            XmlNode newNode3 = doc3.DocumentElement;
                //            doc3.AppendChild(newNode3);
                //            XElement xmlabs3 = XElement.Parse(doc3.OuterXml);
                //            string CounselTypeKind = "";
                //            foreach (XElement abs in xmlabs3.Elements("Item"))
                //            {
                //                CounselTypeKind += abs.Attribute("name").Value;
                //                if (abs.Attribute("name").Value == "其他")
                //                {
                //                    if (abs.Attribute("remark") != null)
                //                        CounselTypeKind += ":" + abs.Attribute("remark").Value;
                //                }
                //                if (abs != xmlabs3.LastNode)
                //                {
                //                    CounselTypeKind += "、";
                //                }
                //            }

                //            //案件類別
                //            cs[row_index, 16].PutValue(CounselTypeKind);

                //            //參與人員(XML)
                //            //cs[row_index, 14].PutValue(InterviewRecord.Attendees);
                //            //處理方式(XML)
                //            //cs[row_index, 15].PutValue(InterviewRecord.CounselType);
                //            //案件類別(XML)
                //            //cs[row_index, 16].PutValue(InterviewRecord.CounselTypeKind);

                //            //內容要點
                //            cs[row_index, 17].PutValue(InterviewRecord.ContentDigest);
                //            //記錄人的登入帳號
                //            cs[row_index, 18].PutValue(InterviewRecord.AuthorID);
                //            //記錄人的姓名
                //            cs[row_index, 19].PutValue(InterviewRecord.AuthorName);
                //            //是否公開(認輔老師、班導師之間互相可見)
                //            cs[row_index, 20].PutValue(InterviewRecord.isPublic? "是":"否");
                //            //記錄人
                //            cs[row_index, 21].PutValue(InterviewRecord.authorRole);                       
                //            row_index++;

                //        }


                //    }
                //    else 
                //    {
                //        if (studentRecord_Dict.ContainsKey(student_id))
                //        {
                //            stuRec = studentRecord_Dict[student_id];
                //        }

                //    }          
                //} 
                #endregion

                printWorker.ReportProgress(100);
            };
            printWorker.RunWorkerCompleted += delegate
            {
                // 以後記得存Excel 都用新版的Xlsx，可以避免ㄧ些不必要的問題(EX: sheet 只能到1023張)
                SaveFileDialog save = new SaveFileDialog();
                save.Title = "另存新檔";
                save.FileName = "輔導案量統計";
                save.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";

                if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        wb.Save(save.FileName, FileFormatType.Excel2003);
                        System.Diagnostics.Process.Start(save.FileName);


                    }
                    catch
                    {
                        MessageBox.Show("檔案儲存失敗");


                    }
                }
            };
            printWorker.ProgressChanged += delegate (object s, ProgressChangedEventArgs e2)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("匯出輔導案量統計", e2.ProgressPercentage);
            };
            printWorker.RunWorkerAsync();

            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //取得所有輔導紀錄中 到底有幾人(因為同一人可以重覆) List<DAO.UDT_CounselStudentInterviewRecordDef> list
        private int GetCounselPeople(String teacherID,List<DAO.UDT_CounselStudentInterviewRecordDef> list)
        {
            int i = 0;
                        
            // 以學生ID 辨識人
            foreach (DAO.UDT_CounselStudentInterviewRecordDef record in list)
            {
                if (!teacher_studentList_Dict.ContainsKey(teacherID))
                {
                    teacher_studentList_Dict.Add(teacherID, new List<int>());

                    if (!teacher_studentList_Dict[teacherID].Contains(record.StudentID))
                    {
                        teacher_studentList_Dict[teacherID].Add(record.StudentID);
                    }
                }
                else
                {
                    if (!teacher_studentList_Dict[teacherID].Contains(record.StudentID))
                    {
                        teacher_studentList_Dict[teacherID].Add(record.StudentID);
                    }
                }                
            }

            i = teacher_studentList_Dict[teacherID].Count();

            return i;
        }

        //多載
        //取得所有輔導紀錄中 到底有幾人(因為同一人可以重覆) List<DAO.UDT_Counsel_home_visit_RecordDef> 
        private int GetCounselPeople(String teacherID, List<DAO.UDT_Counsel_home_visit_RecordDef> list)
        {
            int i = 0;

            // 以學生ID 辨識人
            foreach (DAO.UDT_Counsel_home_visit_RecordDef record in list)
            {
                if (!teacher_studentList_Dict.ContainsKey(teacherID))
                {
                    teacher_studentList_Dict.Add(teacherID, new List<int>());

                    if (!teacher_studentList_Dict[teacherID].Contains(record.StudentID))
                    {
                        teacher_studentList_Dict[teacherID].Add(record.StudentID);
                    }
                }
                else
                {
                    if (!teacher_studentList_Dict[teacherID].Contains(record.StudentID))
                    {
                        teacher_studentList_Dict[teacherID].Add(record.StudentID);
                    }
                }
            }

            i = teacher_studentList_Dict[teacherID].Count();

            return i;
        }

        //建立人數 總加總公式
        private string GetCounselPeopleTotalSumFormula(List<string> departmentList, int teacherCouuter)
        {
            //string formula = "=SUM(D10,D8,D6,D4)";
            string formula = "=SUM(";

            string letter = colLetter[teacherCouuter];

            for (int i = 0; i < departmentList.Count(); i++)
            {
                formula += letter + (4 + i * 2) + ",";
            }

            // 扣掉最後的 , 號
            formula = formula.Substring(0, formula.Length - 1);

            formula += ")";

            return formula;
        }

        //建立人次 總加總公式
        private string GetCounselPeopleCountTotalSumFormula(List<string> departmentList, int teacherCouuter)
        {
            //string formula = "=SUM(D10,D8,D6,D4)";
            string formula = "=SUM(";

            string letter = colLetter[teacherCouuter];

            for (int i = 0; i < departmentList.Count(); i++)
            {
                formula += letter + (5 + i * 2) + ",";
            }

            // 扣掉最後的 , 號
            formula = formula.Substring(0, formula.Length - 1);

            formula += ")";

            return formula;
        }





    }
}
