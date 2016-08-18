using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using System.Xml.Linq;
using FISCA.Data;
using System.Data;
using FISCA.DSAClient;
using FISCA.DSAUtil;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// UDT 資料交換者
    /// </summary>
    public class UDTTransfer
    {
        /// <summary>
        /// 新增測驗資料
        /// </summary>
        /// <param name="QuizData"></param>
        public void InsertQuizData(UDT_QuizDef QuizData)
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_QuizDef> insertList = new List<UDT_QuizDef>();
            insertList.Add(QuizData);
            accHelper.InsertValues(insertList.ToArray());
        }

        /// <summary>
        /// 更新測驗資料
        /// </summary>
        /// <param name="QuizData"></param>
        public void UpdateQuizData(UDT_QuizDef QuizData)
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_QuizDef> UpdateList = new List<UDT_QuizDef>();
            UpdateList.Add(QuizData);
            accHelper.UpdateValues(UpdateList.ToArray());
        }

        /// <summary>
        /// 取得所有測驗
        /// </summary>
        /// <returns></returns>
        public List<UDT_QuizDef> GetAllQuizData()
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_QuizDef> retVal = accHelper.Select<UDT_QuizDef>();
            if (retVal == null)
                retVal = new List<UDT_QuizDef>();

            return retVal;
        }

        /// <summary>
        /// 測過測驗ID取得測驗
        /// </summary>
        /// <returns></returns>
        public UDT_QuizDef GetQuizDataByID(string UID)
        {
            AccessHelper accHelper = new AccessHelper();
            UDT_QuizDef retVal= new UDT_QuizDef();

            if (string.IsNullOrEmpty(UID))
                return retVal;

            string qry = "UID='" + UID + "'";
            List<UDT_QuizDef> retValList = accHelper.Select<UDT_QuizDef>(qry);
            if (retValList.Count > 0)
                retVal = retValList[0];

            return retVal;
        }

        /// <summary>
        /// 刪除測驗資料
        /// </summary>
        /// <param name="QuizName"></param>
        public void DeleteQuizData(UDT_QuizDef QuizData)
        {
            AccessHelper accHelper = new AccessHelper();
            QuizData.Deleted = true;
            List<UDT_QuizDef> delList = new List<UDT_QuizDef>();
            delList.Add(QuizData);
            accHelper.DeletedValues(delList.ToArray());
        }


        /// <summary>
        /// 新增學生測驗資料(單筆)
        /// </summary>
        /// <param name="studQuizData"></param>
        public void InsertStudQuizData(UDT_StudQuizDataDef StudQuizData)
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_StudQuizDataDef> StudQuizDataDefList = new List<UDT_StudQuizDataDef>();

            StudQuizDataDefList.Add(StudQuizData);
            accHelper.InsertValues(StudQuizDataDefList.ToArray());
        }

        /// <summary>
        /// 更新學生測驗資料(單筆)
        /// </summary>
        /// <param name="studQuizData"></param>
        public void UpdateStudQuizData(UDT_StudQuizDataDef StudQuizData)
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_StudQuizDataDef> StudQuizDataDefList = new List<UDT_StudQuizDataDef>();

            StudQuizDataDefList.Add(StudQuizData);
            accHelper.UpdateValues(StudQuizDataDefList.ToArray());
        }

        /// <summary>
        /// 新增多筆學生測驗
        /// </summary>
        /// <param name="studQuizDataList"></param>
        /// <param name="isInsert"></param>
        public void InsertStudQuizDataLlist(List<UDT_StudQuizDataDef> studQuizDataList)
        {
            AccessHelper accHelper = new AccessHelper();
            accHelper.InsertValues(studQuizDataList.ToArray());
        }

        /// <summary>
        /// 更新多筆學生測驗
        /// </summary>
        /// <param name="studQuizDataList"></param>
        /// <param name="isInsert"></param>
        public void UpdateStudQuizDataLlist(List<UDT_StudQuizDataDef> studQuizDataList)
        {
            AccessHelper accHelper = new AccessHelper();
            accHelper.UpdateValues(studQuizDataList.ToArray());
        }


        /// <summary>
        ///  刪除多筆學生測驗
        /// </summary>
        /// <param name="studQuizDataList"></param>
        /// <param name="isInsert"></param>
        public void DeleteStudQuizDataLlist(List<UDT_StudQuizDataDef> studQuizDataList)
        {
            AccessHelper accHelper = new AccessHelper();
            foreach (UDT_StudQuizDataDef rec in studQuizDataList)
                rec.Deleted = true;

            accHelper.DeletedValues(studQuizDataList.ToArray());
        }

        /// <summary>
        /// 依學生編號取得學生測驗資料內容
        /// </summary>
        /// <param name="StudentID"></param>
        /// <returns></returns>
        public List<UDT_StudQuizDataDef> GetStudQuizDataByStudentID(string StudentID)
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_StudQuizDataDef> retVal = new List<UDT_StudQuizDataDef>();

            if (string.IsNullOrEmpty(StudentID))
                return retVal;

            string qry = "ref_student_id='" + StudentID + "'";
            retVal= accHelper.Select<UDT_StudQuizDataDef>(qry);


            return retVal;
        }
        /// <summary>
        /// 依學生編號取得學生測驗資料內容
        /// </summary>
        /// <param name="StudentID"></param>
        /// <returns></returns>
        public List<UDT_StudQuizDataDef> GetStudQuizDataByStudentIDList(List<string> StudentIDList)
        {
            AccessHelper accHelper = new AccessHelper();

            if (StudentIDList.Count == 0)
                return new List<UDT_StudQuizDataDef>();

            string qry = "ref_student_id in("+string.Join (",",StudentIDList.ToArray ())+")";
            Dictionary<int, string> studStatusDict = GetStudentStatusByIDList(StudentIDList);

            List<UDT_StudQuizDataDef> dataList = accHelper.Select<UDT_StudQuizDataDef>(qry);
            if (dataList == null)
                dataList = new List<UDT_StudQuizDataDef>();
            else
            {
                // 填入學生狀態
                foreach (UDT_StudQuizDataDef rec in dataList)
                {
                    if (studStatusDict.ContainsKey(rec.StudentID))
                        rec.StudentStatus = studStatusDict[rec.StudentID];
                }
            }
            return dataList;
        }


        /// <summary>
        /// 依測驗編號取得學生測驗資料內容
        /// </summary>
        /// <param name="StudentID"></param>
        /// <returns></returns>
        public List<UDT_StudQuizDataDef> GetStudQuizDataByQuizID(string QuizID)
        {
            if (string.IsNullOrEmpty(QuizID))
                return new List<UDT_StudQuizDataDef>();

            AccessHelper accHelper = new AccessHelper();
            List<UDT_StudQuizDataDef> retVal;
            string qry = "ref_quiz_uid='" + QuizID + "'";
            retVal= accHelper.Select<UDT_StudQuizDataDef>(qry);
            if (retVal == null)
                retVal = new List<UDT_StudQuizDataDef>();

            return retVal;
        }


        /// <summary>
        /// 建立使用到的 UDT Table
        /// </summary>
        public static void CreateCounselUDTTable()
        {
            //AccessHelper accHelper = new AccessHelper();
            //string query = "UID='1'";
            // 輔導與認輔學生
            FISCA.UDT.SchemaManager Manager = new SchemaManager(new DSConnection(FISCA.Authentication.DSAServices.DefaultDataSource));

            Manager.SyncSchema(new UDT_CounselStudent_ListDef());
            Manager.SyncSchema(new UDT_QuizDef());
            Manager.SyncSchema(new UDT_StudQuizDataDef());
            Manager.SyncSchema(new UDT_CounselStudentInterviewRecordDef());
            Manager.SyncSchema(new UDT_CounselCaseMeetingRecordDef());
            Manager.SyncSchema(new UDT_CounselCareRecordDef());

            Manager.SyncSchema(new UDTQuestionsDataDef());
            Manager.SyncSchema(new UDTYearlyDataDef());
            Manager.SyncSchema(new UDTSemesterDataDef());
            Manager.SyncSchema(new UDTSingleRecordDef());
            Manager.SyncSchema(new UDTMultipleRecordDef());
            Manager.SyncSchema(new UDTRelativeDef());
            Manager.SyncSchema(new UDTSiblingDef());
            Manager.SyncSchema(new UDTPriorityDataDef());

            //accHelper.Select<UDT_CounselStudent_ListDef>(query);
            // 心理測驗題目
            //accHelper.Select<UDT_QuizDef>(query);
            // 心理測驗答案
            //accHelper.Select<UDT_StudQuizDataDef>(query);
            // 晤談
            //accHelper.Select<UDT_CounselStudentInterviewRecordDef>(query);
            // 個案會議
            //accHelper.Select<UDT_CounselCaseMeetingRecordDef>(query);
            // 優先關懷
            //accHelper.Select<UDT_CounselCareRecordDef>(query);                        

            // 學年(綜合紀錄表)
            //accHelper.Select<UDTYearlyDataDef>(query);
            // 學期(綜合紀錄表)
            //accHelper.Select<UDTSemesterDataDef>(query);
            // 單一(綜合紀錄表)
            //accHelper.Select<UDTSingleRecordDef>(query);
            // 多選(綜合紀錄表)
            //accHelper.Select<UDTMultipleRecordDef>(query);
            // 直系血親(綜合紀錄表)
            //accHelper.Select<UDTRelativeDef>(query);
            // 兄弟姊妹(綜合紀錄表)
            //accHelper.Select<UDTSiblingDef>(query);
            // 優先順序(綜合紀錄表)
            //accHelper.Select<UDTPriorityDataDef>(query);
        }

        /// <summary>
        /// 透過學生ID取得晤談紀錄
        /// </summary>
        /// <param name="StudentID"></param>
        /// <returns></returns>
        public List<DAO.UDT_CounselStudentInterviewRecordDef> GetCounselStudentInterviewRecordByStudentID(string StudentID)
        {
            if (string.IsNullOrEmpty(StudentID))
                return new List<UDT_CounselStudentInterviewRecordDef>();

            AccessHelper accHelper = new AccessHelper();
            string qry = "ref_student_id='" + StudentID + "'";
            List<UDT_CounselStudentInterviewRecordDef> retVal;
            retVal= accHelper.Select<UDT_CounselStudentInterviewRecordDef>(qry);

            if (retVal == null)
                retVal = new List<UDT_CounselStudentInterviewRecordDef>();

            return retVal;
        }

        /// <summary>
        /// 透過多筆學生ID取得晤談紀錄
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public List<DAO.UDT_CounselStudentInterviewRecordDef> GetCounselStudentInterviewRecordByStudentIDList(List<string> StudentIDList)
        {
            if (StudentIDList.Count == 0)
                return new List<UDT_CounselStudentInterviewRecordDef>();

            AccessHelper accHelper = new AccessHelper();
            string qry = "ref_student_id in(" + string.Join(",",StudentIDList.ToArray ())+ ")";
            Dictionary<int, string> studStatusDict = GetStudentStatusByIDList(StudentIDList);
            List<UDT_CounselStudentInterviewRecordDef> dataList = accHelper.Select<UDT_CounselStudentInterviewRecordDef>(qry);
            // 填入學生狀態
            if (dataList == null)
                dataList = new List<UDT_CounselStudentInterviewRecordDef>();
            else
            {
                foreach (UDT_CounselStudentInterviewRecordDef rec in dataList)
                    if (studStatusDict.ContainsKey(rec.StudentID))
                        rec.StudentStatus = studStatusDict[rec.StudentID];
            }
            return dataList;
        }


        /// <summary>
        /// 透過多筆教師ID取得晤談紀錄
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public List<DAO.UDT_CounselStudentInterviewRecordDef> GetCounselStudentInterviewRecordByTeacherIDList(List<string> TeacherIDList)
        {
            if (TeacherIDList.Count == 0)
                return new List<UDT_CounselStudentInterviewRecordDef>();

            AccessHelper accHelper = new AccessHelper();
            string qry = "teacher_id in(" + string.Join(",", TeacherIDList.ToArray()) + ")";            
            List<UDT_CounselStudentInterviewRecordDef> dataList = accHelper.Select<UDT_CounselStudentInterviewRecordDef>(qry);
            
            if (dataList == null)
                dataList = new List<UDT_CounselStudentInterviewRecordDef>();
            
             return dataList;
        }


        /// <summary>
        /// 新增一筆晤談紀錄至UDT
        /// </summary>
        /// <param name="StudentInterviewRecord"></param>
        public void InstallCounselStudentInterviewRecord(DAO.UDT_CounselStudentInterviewRecordDef StudentInterviewRecord)
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_CounselStudentInterviewRecordDef> insertList = new List<UDT_CounselStudentInterviewRecordDef>();
            insertList.Add(StudentInterviewRecord);
            accHelper.InsertValues(insertList.ToArray());
        }

        /// <summary>
        /// 新增多筆晤談紀錄至UDT
        /// </summary>
        /// <param name="StudentInterviewRecordList"></param>
        public void InstallCounselStudentInterviewRecordList(List<DAO.UDT_CounselStudentInterviewRecordDef> StudentInterviewRecordList)
        {
            AccessHelper accHelper = new AccessHelper();
            accHelper.InsertValues(StudentInterviewRecordList);
        }


        /// <summary>
        /// 更新一筆晤談紀錄至UDT
        /// </summary>
        /// <param name="StudentInterviewRecord"></param>
        public void UpdateCounselStudentInterviewRecord(DAO.UDT_CounselStudentInterviewRecordDef StudentInterviewRecord)
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_CounselStudentInterviewRecordDef> updateList = new List<UDT_CounselStudentInterviewRecordDef>();
            updateList.Add(StudentInterviewRecord);
            accHelper.UpdateValues(updateList.ToArray());
        }

        /// <summary>
        /// 更新多筆晤談紀錄至UDT
        /// </summary>
        /// <param name="StudentInterviewRecordList"></param>
        public void UpdateCounselStudentInterviewRecordList(List<DAO.UDT_CounselStudentInterviewRecordDef> StudentInterviewRecordList)
        {
            AccessHelper accHelper = new AccessHelper();
            accHelper.UpdateValues(StudentInterviewRecordList);
        }

        /// <summary>
        /// 刪除一筆 UDT 晤談紀錄
        /// </summary>
        /// <param name="StudentInterviewRecord"></param>
        public void DeleteCounselStudentInterviewRecord(DAO.UDT_CounselStudentInterviewRecordDef StudentInterviewRecord)
        {
            StudentInterviewRecord.Deleted = true;
            AccessHelper accHelper = new AccessHelper();
            List<UDT_CounselStudentInterviewRecordDef> delList = new List<UDT_CounselStudentInterviewRecordDef>();
            delList.Add(StudentInterviewRecord);
            accHelper.DeletedValues(delList.ToArray());
        }


        /// <summary>
        /// 新增單筆關懷
        /// </summary>
        /// <param name="CareRecordDef"></param>
        public void InsertCareRecord(UDT_CounselCareRecordDef CareRecordDef)
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_CounselCareRecordDef> CareRecordDefList = new List<UDT_CounselCareRecordDef>();
            CareRecordDefList.Add(CareRecordDef);
            accHelper.InsertValues(CareRecordDefList.ToArray());
        }

        /// <summary>
        /// 新增多筆關懷
        /// </summary>
        /// <param name="CareRecordList"></param>
        public void InsertCareRecordList(List<UDT_CounselCareRecordDef> CareRecordList)
        {
            AccessHelper accHelper = new AccessHelper();
            accHelper.InsertValues(CareRecordList);
        }

        /// <summary>
        /// 更新單筆關懷
        /// </summary>
        /// <param name="CareRecordDef"></param>
        public void UpdateCareRecord(UDT_CounselCareRecordDef CareRecordDef)
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_CounselCareRecordDef> CareRecordDefList = new List<UDT_CounselCareRecordDef>();
            CareRecordDefList.Add(CareRecordDef);
            accHelper.UpdateValues(CareRecordDefList.ToArray());
        }

       
        /// <summary>
        /// 更新多筆關懷
        /// </summary>
        /// <param name="CareRecordList"></param>
        public void UpdateCareRecordList(List<UDT_CounselCareRecordDef> CareRecordList)
        {
            AccessHelper accHelper = new AccessHelper();
            accHelper.UpdateValues(CareRecordList);
        }


        /// <summary>
        /// 刪除單筆關懷
        /// </summary>
        /// <param name="CareRecordDef"></param>
        public void DeleteCareRecord(UDT_CounselCareRecordDef CareRecordDef)
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_CounselCareRecordDef> CareRecordDefList = new List<UDT_CounselCareRecordDef>();
            CareRecordDef.Deleted = true;
            CareRecordDefList.Add(CareRecordDef);
            accHelper.DeletedValues(CareRecordDefList.ToArray());

        }

        /// <summary>
        /// 依學生編號取得關懷
        /// </summary>
        /// <param name="StudentID"></param>
        /// <returns></returns>
        public List<UDT_CounselCareRecordDef> GetCareRecordsByStudentID(string StudentID)
        {
            if (string.IsNullOrEmpty(StudentID))
                return new List<UDT_CounselCareRecordDef>();

            AccessHelper accHelper = new AccessHelper();
            List<UDT_CounselCareRecordDef> retVal;
            List<UDT_CounselCareRecordDef> CareRecordList = new List<UDT_CounselCareRecordDef>();
            string qry = "ref_student_id='" + StudentID + "'";
            retVal= accHelper.Select<UDT_CounselCareRecordDef>(qry);
            if (retVal == null)
                retVal = new List<UDT_CounselCareRecordDef>();

            return retVal;
        }

        /// <summary>
        /// 依多筆學生編號取得關懷
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public List<UDT_CounselCareRecordDef> GetCareRecordsByStudentIDList(List<string> StudentIDList)
        {
            if (StudentIDList.Count == 0)
                return new List<UDT_CounselCareRecordDef>();

            AccessHelper accHelper = new AccessHelper();
            string qry = "ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
            List<UDT_CounselCareRecordDef> dataList =accHelper.Select<UDT_CounselCareRecordDef>(qry);
            Dictionary<int, string> studStatusDict = GetStudentStatusByIDList(StudentIDList);

            if (dataList == null)
                dataList = new List<UDT_CounselCareRecordDef>();
            else
            {
                // 填入學生狀態
                foreach (UDT_CounselCareRecordDef rec in dataList)
                    if (studStatusDict.ContainsKey(rec.StudentID))
                        rec.StudentStatus = studStatusDict[rec.StudentID];
            }
            return dataList;
        }

        



        /// <summary>
        /// 新增單筆個案會議
        /// </summary>
        /// <param name="CaseMeetingRecord"></param>
        public void InsertCaseMeetingRecord(DAO.UDT_CounselCaseMeetingRecordDef CaseMeetingRecord)
        {
            List<DAO.UDT_CounselCaseMeetingRecordDef> InsertCaseMeetingRecordList = new List<UDT_CounselCaseMeetingRecordDef>();
            InsertCaseMeetingRecordList.Add(CaseMeetingRecord);
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.InsertValues(InsertCaseMeetingRecordList.ToArray());
        }

        /// <summary>
        /// 新增多筆個案會議
        /// </summary>
        /// <param name="CaseMeetingRecordList"></param>
        public void InsertCaseMeetingRecordList(List<DAO.UDT_CounselCaseMeetingRecordDef> CaseMeetingRecordList)
        {
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.InsertValues(CaseMeetingRecordList);
        }


        /// <summary>
        /// 更新單筆個案會議
        /// </summary>
        /// <param name="CaseMeetingRecord"></param>
        public void UpdateCaseMeetingRecord(DAO.UDT_CounselCaseMeetingRecordDef CaseMeetingRecord)
        {
            List<DAO.UDT_CounselCaseMeetingRecordDef> UpdateCaseMeetingRecordList = new List<UDT_CounselCaseMeetingRecordDef>();
            UpdateCaseMeetingRecordList.Add(CaseMeetingRecord);
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.UpdateValues(UpdateCaseMeetingRecordList.ToArray());

        }

        /// <summary>
        /// 更新多筆個案會議
        /// </summary>
        /// <param name="CaseMeetingRecordList"></param>
        public void UpdateCaseMeetingRecordList(List<DAO.UDT_CounselCaseMeetingRecordDef> CaseMeetingRecordList)
        {
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.UpdateValues(CaseMeetingRecordList);
        }

        /// <summary>
        /// 刪除1筆個案會議
        /// </summary>
        /// <param name="CaseMeetingRecord"></param>
        public void DeleteCaseMeetingRecord(DAO.UDT_CounselCaseMeetingRecordDef CaseMeetingRecord)
        {
            CaseMeetingRecord.Deleted = true;
            List<DAO.UDT_CounselCaseMeetingRecordDef> DelCaseMeetingRecordList = new List<UDT_CounselCaseMeetingRecordDef>();
            DelCaseMeetingRecordList.Add(CaseMeetingRecord);
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.DeletedValues(DelCaseMeetingRecordList.ToArray());

        }

        /// <summary>
        /// 透過學生ID取得學生個案會議
        /// </summary>
        /// <param name="StudentID"></param>
        /// <returns></returns>s
        public List<DAO.UDT_CounselCaseMeetingRecordDef> GetCaseMeetingRecordListByStudentID(string StudentID)
        {
            if (string.IsNullOrEmpty(StudentID))
                return new List<UDT_CounselCaseMeetingRecordDef>();

            string qry = "ref_student_id='" + StudentID + "'";
            AccessHelper accessHelper = new AccessHelper();
            List<UDT_CounselCaseMeetingRecordDef> retVal;
            retVal= accessHelper.Select<UDT_CounselCaseMeetingRecordDef>(qry);

            if (retVal == null)
                retVal = new List<UDT_CounselCaseMeetingRecordDef>();

            return retVal;
        }

        /// <summary>
        /// 透過多筆學生ID取得學生個案會議
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public List<DAO.UDT_CounselCaseMeetingRecordDef> GetCaseMeetingRecordListByStudentIDList(List<string> StudentIDList)
        {
            if (StudentIDList.Count == 0)
                return new List<UDT_CounselCaseMeetingRecordDef>();
   
            AccessHelper accessHelper = new AccessHelper();
            string qry = "ref_student_id in(" + string.Join(",",StudentIDList.ToArray ()) + ")";
            List<UDT_CounselCaseMeetingRecordDef> dataList= accessHelper.Select<UDT_CounselCaseMeetingRecordDef>(qry);
            Dictionary<int, string> studStatusDict = GetStudentStatusByIDList(StudentIDList);
            if (dataList == null)
                dataList = new List<UDT_CounselCaseMeetingRecordDef>();
            else
            {
                // 填入學生狀態
                foreach (UDT_CounselCaseMeetingRecordDef rec in dataList)
                    if (studStatusDict.ContainsKey(rec.StudentID))
                        rec.StudentStatus = studStatusDict[rec.StudentID];
            }
            return dataList;
        }


        /// <summary>
        /// 透過學生IDList取得綜合表現答案
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<DAO.UDT_ABCardDataDef> GetABCardDataListByStudentList(List<string> StudentIDList)
        {
            List<DAO.UDT_ABCardDataDef> retVal = new List<UDT_ABCardDataDef>();
            if (StudentIDList.Count > 0)
            {
                AccessHelper accessHelper = new AccessHelper();
                string qry = "ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                retVal = accessHelper.Select<UDT_ABCardDataDef>(qry);
            }
            return retVal;
        }
        
        /// <summary>
        /// 透過學生ID取得學生輔導老師
        /// </summary>
        /// <param name="StudentID"></param>
        /// <returns></returns>
        public List<DAO.UDT_CounselStudent_ListDef> GetCounselStudentListID(string StudentID)
        {
            if (string.IsNullOrEmpty(StudentID))
                return new List<UDT_CounselStudent_ListDef>();

            string qry = "ref_student_id='" + StudentID + "'";
            AccessHelper accessHelper = new AccessHelper();
            return accessHelper.Select<UDT_CounselStudent_ListDef>(qry);        
        }

        /// <summary>
        /// 透過TeacherTagID取得學生輔導老師
        /// </summary>
        /// <param name="StudentID"></param>
        /// <returns></returns>
        public List<DAO.UDT_CounselStudent_ListDef> GetCounselStudentListByTeacherTagID(int TeacherTagID)
        {
            if (TeacherTagID == 0)
                return new List<UDT_CounselStudent_ListDef>();

            string qry = "ref_teacher_tag_id='" + TeacherTagID + "'";
            AccessHelper accessHelper = new AccessHelper();
            return accessHelper.Select<UDT_CounselStudent_ListDef>(qry);
        }


        /// <summary>
        /// 透過教師ID,類別名稱取得所TagID
        /// </summary>
        /// <param name="TeacherID"></param>
        /// <param name="TagName"></param>
        /// <returns></returns>
        public int? GetTeacherTagIDByTeacherID(string TeacherID,string TagName)
        {
            if(string.IsNullOrEmpty(TeacherID) || string.IsNullOrEmpty(TagName ))
            return null ;
            
            string strSQL="select tag_teacher.id from tag_teacher inner join tag on tag_teacher.ref_tag_id=tag.id where ref_teacher_id="+TeacherID+" and tag.name='"+TagName+"';";
            int TeacherTagID=0; ;
            QueryHelper helper = new QueryHelper();
            DataTable dt= helper.Select(strSQL);

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)               
                    int.TryParse(dr[0].ToString(), out TeacherTagID);                            
            }
            if (TeacherTagID == 0)
                return null;
            else
                return TeacherTagID;
        }

        /// <summary>
        /// 透過多筆學生ID，取得學生輔導老師
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public List<DAO.UDT_CounselStudent_ListDef> GetCounselStudentListIDs(List<string> StudentIDList)
        {
            if (StudentIDList.Count == 0)
                return new List<UDT_CounselStudent_ListDef>();

            AccessHelper accHelper = new AccessHelper();

            string qry = "ref_student_id in(" + string.Join(",",StudentIDList.ToArray ()) + ")";
            return accHelper.Select<UDT_CounselStudent_ListDef>(qry);
        }

        /// <summary>
        /// 新增多筆學生輔導老師
        /// </summary>
        /// <param name="CounselStudent_List"></param>
        public void InsertCounselStudentList(List<DAO.UDT_CounselStudent_ListDef> CounselStudent_List)
        {
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.InsertValues(CounselStudent_List.ToArray());
        }

        /// <summary>
        /// 更新多筆學生輔導老師
        /// </summary>
        /// <param name="CounselStudent_List"></param>
        public void UpdateCounselStudentList(List<DAO.UDT_CounselStudent_ListDef> CounselStudent_List)
        {
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.UpdateValues(CounselStudent_List.ToArray());
        }


        /// <summary>
        /// 刪除多筆學生輔導老師
        /// </summary>
        /// <param name="CounselStudent_List"></param>
        public void DeleteCounselStudentList(List<DAO.UDT_CounselStudent_ListDef> CounselStudent_List)
        {
            List<UDT_CounselStudent_ListDef> DelCounselStudentList = new List<UDT_CounselStudent_ListDef>();
            foreach (UDT_CounselStudent_ListDef rec in CounselStudent_List)
            {
                rec.Deleted = true;
                DelCounselStudentList.Add(rec);
            }
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.DeletedValues(DelCounselStudentList.ToArray());
        }

        /// <summary>
        /// 新增多筆輔導使用者自訂欄位
        /// </summary>
        /// <param name="dataList"></param>
        public void InsertCounselUsereDefinfDataList(List<UDT_CounselUserDefDataDef> dataList)
        {
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.InsertValues(dataList);        
        }


        /// <summary>
        /// 更新多筆輔導使用者自訂欄位
        /// </summary>
        /// <param name="dataList"></param>
        public void UpdateCounselUserDefineDataList(List<UDT_CounselUserDefDataDef> dataList)
        {
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.UpdateValues(dataList);        
        }

        /// <summary>
        /// 刪除多筆輔導使用者自訂欄位
        /// </summary>
        /// <param name="dataList"></param>
        public void DeleteCounselUserDefineDataList(List<UDT_CounselUserDefDataDef> dataList)
        {
            List<UDT_CounselUserDefDataDef> dtList = new List<UDT_CounselUserDefDataDef>();
            foreach (UDT_CounselUserDefDataDef rec in dataList)
            {
                rec.Deleted = true;
                dtList.Add(rec);
            }
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.DeletedValues(dtList);
        }

        /// <summary>
        /// 依學生IDList取得輔導自訂欄位資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public List<UDT_CounselUserDefDataDef> GetCounselUserDefineDataByStudentIDList(List<string> StudentIDList)
        {
            if (StudentIDList.Count == 0)
                return new List<UDT_CounselUserDefDataDef>();

            AccessHelper accessHelper = new AccessHelper();
            Dictionary<int, string> studStatusDict = GetStudentStatusByIDList(StudentIDList);

            string qry = "ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
            List<UDT_CounselUserDefDataDef> dataList=accessHelper.Select<UDT_CounselUserDefDataDef>(qry);
            if (dataList == null)
                dataList = new List<UDT_CounselUserDefDataDef>();
            else
            {
                // 填入學生狀態
                foreach (UDT_CounselUserDefDataDef rec in dataList)
                    if (studStatusDict.ContainsKey(rec.StudentID))
                        rec.StudentStatus = studStatusDict[rec.StudentID];
            }
            return dataList;
        }

        /// <summary>
        /// 透過學生ID取得學生狀態ID:int,
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<int, string> GetStudentStatusByIDList(List<string> StudentIDList)
        {
            if (StudentIDList.Count==0)
                return new Dictionary<int, string>();

            Dictionary<int, string> retVal = new Dictionary<int, string>();
            QueryHelper qh = new QueryHelper();
            string strSQL = "select id,case when status=1 then '一般' when status=2 then '延修' when status=4 then '休學' when status=8 then '輟學' when status=16 then '畢業或離校' when status=256 then '刪除' else '' end from student;";
            DataTable dt= qh.Select(strSQL);
            foreach(DataRow dr in dt.Rows)
            {
                int id = int.Parse(dr[0].ToString());
                if (!retVal.ContainsKey(id))
                    retVal.Add(id, dr[1].ToString());
            }
            return retVal;
        }

        /// <summary>
        /// 取得ABCard樣板內容
        /// </summary>
        /// <returns></returns>
        public static List<UDT_ABCardTemplateDefinitionDef> GetABCardTemplate()
        {
            AccessHelper accHelper = new AccessHelper();
            List<UDT_ABCardTemplateDefinitionDef> retVal;
            retVal= accHelper.Select<UDT_ABCardTemplateDefinitionDef>();
            if (retVal == null)
                retVal = new List<UDT_ABCardTemplateDefinitionDef>();

            return retVal;
        }

        /// <summary>
        /// 新增ABCard樣板內容
        /// </summary>
        /// <returns></returns>
        public static void InsertABCardTemplate(List<UDT_ABCardTemplateDefinitionDef> RecList)
        {
            AccessHelper accHelper = new AccessHelper();
            accHelper.InsertValues(RecList);
        }

        /// <summary>
        /// 更新ABCard樣板內容
        /// </summary>
        /// <param name="RecList"></param>
        public static void UpdateABCardTemplate(List<UDT_ABCardTemplateDefinitionDef> RecList)
        {
            AccessHelper accHelper = new AccessHelper();
            accHelper.UpdateValues(RecList);
        }


        /// <summary>
        /// 刪除ABCard樣板內容
        /// </summary>
        /// <param name="RecList"></param>
        public static void DeleteABCardTemplate(List<UDT_ABCardTemplateDefinitionDef> RecList)
        {
            AccessHelper accHelper = new AccessHelper();
            foreach (UDT_ABCardTemplateDefinitionDef rec in RecList)
                rec.Deleted = true;

            accHelper.DeletedValues(RecList);
        }

        /// <summary>
        /// 新增輔導系統設定內容
        /// </summary>
        /// <param name="RecList"></param>
        public static void InsertSystemList(List<UDT_SystemListDef> RecList)
        {
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.InsertValues(RecList);        
        }

        /// <summary>
        /// 修改輔導系統設定內容
        /// </summary>
        /// <param name="RecList"></param>
        public static void UpdateSystemList(List<UDT_SystemListDef> RecList)
        {
            AccessHelper accessHelper = new AccessHelper();
            accessHelper.UpdateValues(RecList);            
        }

        /// <summary>
        /// 透過名稱取得設定內容
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static UDT_SystemListDef GetSystemListByName(string name)
        {
            UDT_SystemListDef retVal = new UDT_SystemListDef();
            string qry = "name='" + name + "'";
            AccessHelper accessHelper = new AccessHelper();
            List<UDT_SystemListDef> dataList = accessHelper.Select<UDT_SystemListDef>(qry);

            if (dataList.Count > 0)
                retVal = dataList[0];

            return retVal;
        }

        /// <summary>
        /// 刪除輔導系統設定內容
        /// </summary>
        /// <param name="RecList"></param>
        public static void DeleteSystemList(List<UDT_SystemListDef> RecList)
        {
            AccessHelper accessHelper = new AccessHelper();
            foreach (UDT_SystemListDef rec in RecList)
                rec.Deleted = true;

            accessHelper.DeletedValues(RecList);        
        }

        /// <summary>
        /// 新增綜合紀錄表 親屬資訊
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTRelativeInsert(List<UDTRelativeDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.InsertValues(dataList);
            }        
        }

        /// <summary>
        /// 更新綜合紀錄表 親屬資訊
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTRelativeUpdate(List<UDTRelativeDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.UpdateValues(dataList);
            }
        }

        /// <summary>
        /// 刪除綜合紀錄表 親屬資訊
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTRelativeDelete(List<UDTRelativeDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                foreach (UDTRelativeDef data in dataList)
                    data.Deleted = true;

                accHelper.DeletedValues(dataList);
            }
        }

        /// <summary>
        /// 依學生ID 查詢 綜合紀錄表 親屬資訊 
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<UDTRelativeDef> ABUDTRelativeSelectByStudentIDList(List<string> StudentIDList)
        {
            List<UDTRelativeDef> retVal = new List<UDTRelativeDef>();

            if (StudentIDList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                string qry = "ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                retVal = accHelper.Select<UDTRelativeDef>(qry);                
            }

            return retVal;
        }

        /// <summary>
        /// 新增綜合紀錄表 兄弟姊妹資訊
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTSiblingInsert(List<UDTSiblingDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.InsertValues(dataList);
            }
        }

        /// <summary>
        /// 更新綜合紀錄表 兄弟姊妹資訊
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTSiblingUpdate(List<UDTSiblingDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.UpdateValues(dataList);
            }
        }

        /// <summary>
        /// 刪除綜合紀錄表 兄弟姊妹資訊
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTSiblingDelete(List<UDTSiblingDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                foreach (UDTSiblingDef data in dataList)
                    data.Deleted = true;

                accHelper.DeletedValues(dataList);
            }
        }

        /// <summary>
        /// 依學生ID 查詢 綜合紀錄表 兄弟姊妹資訊 
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<UDTSiblingDef> ABUDTSiblingSelectByStudentIDList(List<string> StudentIDList)
        {
            List<UDTSiblingDef> retVal = new List<UDTSiblingDef>();

            if (StudentIDList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                string qry = "ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                retVal = accHelper.Select<UDTSiblingDef>(qry);
            }

            return retVal;
        }

        /// <summary>
        /// 新增綜合紀錄表 單一記錄
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTSingleRecordInsert(List<UDTSingleRecordDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.InsertValues(dataList);
            }
        }

        /// <summary>
        /// 更新綜合紀錄表 單一記錄
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTSingleRecordUpdate(List<UDTSingleRecordDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.UpdateValues(dataList);
            }
        }

        /// <summary>
        /// 刪除綜合紀錄表 單一記錄
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTSingleRecordDelete(List<UDTSingleRecordDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                foreach (UDTSingleRecordDef data in dataList)
                    data.Deleted = true;

                accHelper.DeletedValues(dataList);
            }
        }

        /// <summary>
        /// 依學生ID 查詢 綜合紀錄表 單一記錄 
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<UDTSingleRecordDef> ABUDTSingleRecordSelectByStudentIDList(List<string> StudentIDList)
        {
            List<UDTSingleRecordDef> retVal = new List<UDTSingleRecordDef>();

            if (StudentIDList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                string qry = "ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                retVal = accHelper.Select<UDTSingleRecordDef>(qry);
            }

            return retVal;
        }


        /// <summary>
        /// 新增綜合紀錄表 複選記錄
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTMultipleRecordInsert(List<UDTMultipleRecordDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.InsertValues(dataList);
            }
        }

        /// <summary>
        /// 更新綜合紀錄表 複選記錄
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTMultipleRecordUpdate(List<UDTMultipleRecordDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.UpdateValues(dataList);
            }
        }

        /// <summary>
        /// 刪除綜合紀錄表 複選記錄
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTMultipleRecordDelete(List<UDTMultipleRecordDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                foreach (UDTMultipleRecordDef data in dataList)
                    data.Deleted = true;

                accHelper.DeletedValues(dataList);
            }
        }

        /// <summary>
        /// 依學生ID 查詢 綜合紀錄表 複選記錄 
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<UDTMultipleRecordDef> ABUDTMultipleRecordSelectByStudentIDList(List<string> StudentIDList)
        {
            List<UDTMultipleRecordDef> retVal = new List<UDTMultipleRecordDef>();

            if (StudentIDList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                string qry = "ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                retVal = accHelper.Select<UDTMultipleRecordDef>(qry);
            }

            return retVal;
        }

        /// <summary>
        /// 新增綜合紀錄表 每年資料
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTYearlyDataInsert(List<UDTYearlyDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.InsertValues(dataList);
            }
        }

        /// <summary>
        /// 更新綜合紀錄表 每年資料
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTYearlyDataUpdate(List<UDTYearlyDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.UpdateValues(dataList);
            }
        }

        /// <summary>
        /// 刪除綜合紀錄表 每年資料
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTYearlyDataDelete(List<UDTYearlyDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                foreach (UDTYearlyDataDef data in dataList)
                    data.Deleted = true;

                accHelper.DeletedValues(dataList);
            }
        }

        /// <summary>
        /// 依學生ID 查詢 綜合紀錄表 每年資料 
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<UDTYearlyDataDef> ABUDTYearlyDataSelectByStudentIDList(List<string> StudentIDList)
        {
            List<UDTYearlyDataDef> retVal = new List<UDTYearlyDataDef>();

            if (StudentIDList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                string qry = "ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                retVal = accHelper.Select<UDTYearlyDataDef>(qry);
            }

            return retVal;
        }

        /// <summary>
        /// 新增綜合紀錄表 每學期資料
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTSemesterDataInsert(List<UDTSemesterDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.InsertValues(dataList);
            }
        }

        /// <summary>
        /// 更新綜合紀錄表 每學期資料
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTSemesterDataUpdate(List<UDTSemesterDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.UpdateValues(dataList);
            }
        }

        /// <summary>
        /// 刪除綜合紀錄表 每學期資料
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTSemesterDataDelete(List<UDTSemesterDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                foreach (UDTSemesterDataDef data in dataList)
                    data.Deleted = true;

                accHelper.DeletedValues(dataList);
            }
        }

        /// <summary>
        /// 依學生ID 查詢 綜合紀錄表 每學期資料 
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<UDTSemesterDataDef> ABUDTSemesterDataSelectByStudentIDList(List<string> StudentIDList)
        {
            List<UDTSemesterDataDef> retVal = new List<UDTSemesterDataDef>();

            if (StudentIDList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                string qry = "ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                retVal = accHelper.Select<UDTSemesterDataDef>(qry);
            }

            return retVal;
        }

        /// <summary>
        /// 新增綜合紀錄表 題目資料
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTQuestionsDataInsert(List<UDTQuestionsDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.InsertValues(dataList);
            }
        }

        /// <summary>
        /// 更新綜合紀錄表 題目資料
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTQuestionsDataUpdate(List<UDTQuestionsDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.UpdateValues(dataList);
            }
        }

        /// <summary>
        /// 刪除綜合紀錄表 題目資料
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTQuestionsDataDelete(List<UDTQuestionsDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                foreach (UDTQuestionsDataDef data in dataList)
                    data.Deleted = true;

                accHelper.DeletedValues(dataList);
            }
        }

        /// <summary>
        /// 取得綜合表現記錄所有題目
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<UDTQuestionsDataDef> ABUDTQuestionsDataSelectAll()
        {
            List<UDTQuestionsDataDef> retVal = new List<UDTQuestionsDataDef>();
         
                AccessHelper accHelper = new AccessHelper();
               retVal = accHelper.Select<UDTQuestionsDataDef>();            

            return retVal;
        }

        /// <summary>
        /// 清空綜合表現記錄所有題目
        /// </summary>
        public static void ABUDTQuestionsDataDeleteAll()
        {           
            ABUDTQuestionsDataDelete(ABUDTQuestionsDataSelectAll());
        }

        /// <summary>
        /// 依群組名稱取得綜合表現記錄該群組所有題目
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<UDTQuestionsDataDef> ABUDTQuestionsDataSelectByGroupName(string Name)
        {
            List<UDTQuestionsDataDef> retVal = new List<UDTQuestionsDataDef>();

            AccessHelper accHelper = new AccessHelper();
            string query = "group_name='"+Name+"'";
            retVal = accHelper.Select<UDTQuestionsDataDef>(query);

            return retVal;
        }

        /// <summary>
        /// 新增綜合紀錄表 優先順序
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTPriorityDataInsert(List<UDTPriorityDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.InsertValues(dataList);
            }
        }

        /// <summary>
        /// 更新綜合紀錄表 優先順序
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTPriorityDataUpdate(List<UDTPriorityDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                accHelper.UpdateValues(dataList);
            }
        }

        /// <summary>
        /// 刪除綜合紀錄表 優先順序
        /// </summary>
        /// <param name="dataList"></param>
        public static void ABUDTPriorityDataDelete(List<UDTPriorityDataDef> dataList)
        {
            if (dataList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                foreach (UDTPriorityDataDef data in dataList)
                    data.Deleted = true;

                accHelper.DeletedValues(dataList);
            }
        }

        /// <summary>
        /// 依學生ID 查詢 綜合紀錄表 優先順序 
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<UDTPriorityDataDef> ABUDTPriorityDataSelectByStudentIDList(List<string> StudentIDList)
        {
            List<UDTPriorityDataDef> retVal = new List<UDTPriorityDataDef>();

            if (StudentIDList.Count > 0)
            {
                AccessHelper accHelper = new AccessHelper();
                string qry = "ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                retVal = accHelper.Select<UDTPriorityDataDef>(qry);
            }

            return retVal;
        }


    }
}
