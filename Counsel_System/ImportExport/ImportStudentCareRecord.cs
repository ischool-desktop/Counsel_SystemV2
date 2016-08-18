using System;
using System.Collections.Generic;
using System.Linq;
using Campus.Import;
using Campus.DocumentValidator;
using System.Text;

namespace Counsel_System2.ImportExport
{
    /// <summary>
    /// 匯入優先關懷
    /// </summary>
    public class ImportStudentCareRecord : ImportWizard
    {
        private ImportOption mOption;        
        DAO.UDTTransfer _UDTTransfer;

        public ImportStudentCareRecord()
        {
            this.IsSplit = false;
            this.IsLog = false;
        }

        public override ImportAction GetSupportActions()
        {
            return ImportAction.InsertOrUpdate;
        }

        public override string GetValidateRule()
        {           
            return Properties.Resources.ImportStudentCareRecordVal;
        }

        public override string Import(List<Campus.DocumentValidator.IRowStream> Rows)
        {
            List<DAO.UDT_CounselCareRecordDef> InsertData = new List<DAO.UDT_CounselCareRecordDef>();
            List<DAO.UDT_CounselCareRecordDef> UpdateData = new List<DAO.UDT_CounselCareRecordDef>();
            List<DAO.UDT_CounselCareRecordDef> HasData = new List<DAO.UDT_CounselCareRecordDef>();
            
            // -- 處理 log
            Dictionary<string, StringBuilder> LogData = new Dictionary<string, StringBuilder>();
            // 學生ID List
             List<int> studIdList= new List<int> ();
            foreach (IRowStream ir in Rows)
            {
                int i;
                if (ir.Contains("學號") && ir.Contains("狀態"))
                {
                    if (int.TryParse(Utility.GetStudentID(ir.GetValue("學號"), ir.GetValue("狀態")).ToString(), out i))
                        studIdList.Add(i);
                }
            } 
            // 取得學生名稱log 用
            Dictionary<string, string> StudentNameDict = new Dictionary<string, string>();
            foreach (KeyValuePair<int, string> data in Utility.GetConvertStringDict1fromDB(studIdList))
                StudentNameDict.Add(data.Key.ToString(), data.Value);
            
            DAO.LogTransfer _LogTransfer = new DAO.LogTransfer();

            List<string> StudentIDList = new List<string>();
            // 取得學生狀態對應
            foreach (IRowStream ir in Rows)
            {
                if (ir.Contains("學號") && ir.Contains("狀態"))
                    StudentIDList.Add(Utility.GetStudentID(ir.GetValue("學號"), ir.GetValue("狀態")).ToString());
            }
            // 已有資料
            HasData = _UDTTransfer.GetCareRecordsByStudentIDList(StudentIDList);

            // 取得教師帳號比對用
            Dictionary<string, string> teacherNameLoginIDDict = Utility.GetTeacherNameLoginIDStatus1();

            int TotalCount = 0, NewIdx = 0;

            foreach (IRowStream ir in Rows)
            {
                TotalCount++;
                this.ImportProgress = TotalCount;
                DAO.UDT_CounselCareRecordDef CounselCareRecord = null;
                int sid = 0;
                if (ir.Contains("學號") && ir.Contains("狀態"))
                {
                    string key =ir.GetValue("學號") + "_";
                    if (Global._StudentStatusDBDict.ContainsKey(ir.GetValue("狀態")))
                        sid = Utility.GetStudentID(ir.GetValue("學號"), ir.GetValue("狀態"));

                    DateTime dt;
                    // 當同一位學生有相同會議日期與會議事由，當作是更新，否則新增
                    if (DateTime.TryParse(ir.GetValue("立案日期"), out dt))
                    {
                        foreach (DAO.UDT_CounselCareRecordDef rec in HasData.Where(x => x.StudentID == sid))
                        {
                            if (rec.FileDate.HasValue)
                                if (rec.FileDate.Value.ToShortDateString() == dt.ToShortDateString())
                                    if (rec.CaseCategory == ir.GetValue("個案類別") && rec.CaseOrigin==ir.GetValue("個案來源"))
                                        CounselCareRecord = rec;
                        }
                    }

                    bool isNew = false;
                    if (CounselCareRecord == null)
                    {
                        CounselCareRecord = new DAO.UDT_CounselCareRecordDef();
                        isNew = true;
                        NewIdx++;
                    }
                    string StudID=sid.ToString();
                    string insertKey = "A" + NewIdx;
                    // 學生編號
                    CounselCareRecord.StudentID = sid;
                    // 立案日期                    
                    if (isNew)
                        _LogTransfer.AddBatchInsertLog(StudID, insertKey, "立案日期", dt.ToShortDateString());
                    else
                        _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "立案日期", CounselCareRecord.FileDate.Value.ToShortDateString(), dt.ToShortDateString());
                     CounselCareRecord.FileDate = dt;

                     if (ir.Contains("代號"))
                     {
                         if (isNew)
                             _LogTransfer.AddBatchInsertLog(StudID, insertKey, "代號", ir.GetValue("代號"));
                         else
                             _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "代號", CounselCareRecord.CodeName, ir.GetValue("代號"));
                         CounselCareRecord.CodeName = ir.GetValue("代號");
                     }

                     if (isNew)
                         _LogTransfer.AddBatchInsertLog(StudID, insertKey, "個案類別", ir.GetValue("個案類別"));
                     else
                         _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "個案類別", CounselCareRecord.CaseCategory, ir.GetValue("個案類別"));
                    CounselCareRecord.CaseCategory = ir.GetValue("個案類別");



                    if (ir.Contains("個案類別備註"))
                    {
                        if (isNew)
                            _LogTransfer.AddBatchInsertLog(StudID, insertKey, "個案類別備註", ir.GetValue("個案類別備註"));
                        else
                            _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "個案類別備註", CounselCareRecord.CaseCategoryRemark, ir.GetValue("個案類別備註"));
                        CounselCareRecord.CaseCategoryRemark = ir.GetValue("個案類別備註");
                    }

                    if (isNew)
                        _LogTransfer.AddBatchInsertLog(StudID, insertKey, "個案來源", ir.GetValue("個案來源"));
                    else
                        _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "個案來源", CounselCareRecord.CaseOrigin, ir.GetValue("個案來源"));
                    CounselCareRecord.CaseOrigin = ir.GetValue("個案來源");

                    if (ir.Contains("個案來源備註"))
                    {
                        if (isNew)
                            _LogTransfer.AddBatchInsertLog(StudID, insertKey, "個案來源備註", ir.GetValue("個案來源備註"));
                        else
                            _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "個案來源備註", CounselCareRecord.CaseOriginRemark, ir.GetValue("個案來源備註"));
                        CounselCareRecord.CaseOriginRemark = ir.GetValue("個案來源備註");
                    }
                    if (ir.Contains("優勢能力及財力"))
                    {
                        if (isNew)
                            _LogTransfer.AddBatchInsertLog(StudID, insertKey, "優勢能力及財力", ir.GetValue("優勢能力及財力"));
                        else
                            _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "優勢能力及財力", CounselCareRecord.Superiority, ir.GetValue("優勢能力及財力"));
                        CounselCareRecord.Superiority = ir.GetValue("優勢能力及財力");
                    }

                    if (ir.Contains("弱勢能力及財力"))
                    {
                        if (isNew)
                            _LogTransfer.AddBatchInsertLog(StudID, insertKey, "弱勢能力及財力", ir.GetValue("弱勢能力及財力"));
                        else
                            _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "弱勢能力及財力", CounselCareRecord.Weakness, ir.GetValue("弱勢能力及財力"));
                        CounselCareRecord.Weakness = ir.GetValue("弱勢能力及財力");
                    }
                    if (ir.Contains("輔導人員輔導目標"))
                    {
                        if (isNew)
                            _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導人員輔導目標", ir.GetValue("輔導人員輔導目標"));
                        else
                            _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "輔導人員輔導目標", CounselCareRecord.CounselGoal, ir.GetValue("輔導人員輔導目標"));
                        CounselCareRecord.CounselGoal = ir.GetValue("輔導人員輔導目標");
                    }
                    if (ir.Contains("校外協輔機構"))
                    {
                        if (isNew)
                            _LogTransfer.AddBatchInsertLog(StudID, insertKey, "校外協輔機構", ir.GetValue("校外協輔機構"));
                        else
                            _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "校外協輔機構", CounselCareRecord.OtherInstitute, ir.GetValue("校外協輔機構"));
                        CounselCareRecord.OtherInstitute = ir.GetValue("校外協輔機構");
                    }
                    if (ir.Contains("輔導人員輔導方式"))
                    {
                        if (isNew)
                            _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導人員輔導方式", ir.GetValue("輔導人員輔導方式"));
                        else
                            _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "輔導人員輔導方式", CounselCareRecord.CounselType, ir.GetValue("輔導人員輔導方式"));
                        CounselCareRecord.CounselType = ir.GetValue("輔導人員輔導方式");
                    }
                    if (ir.Contains("協同輔導人員協助導師事項"))
                    {
                        if (isNew)
                            _LogTransfer.AddBatchInsertLog(StudID, insertKey, "協同輔導人員協助導師事項", ir.GetValue("協同輔導人員協助導師事項"));
                        else
                            _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "協同輔導人員協助導師事項", CounselCareRecord.AssistedMatter, ir.GetValue("協同輔導人員協助導師事項"));
                        CounselCareRecord.AssistedMatter = ir.GetValue("協同輔導人員協助導師事項");
                    }

                    if (ir.Contains("記錄者姓名"))
                    {
                        if (isNew)
                            _LogTransfer.AddBatchInsertLog(StudID, insertKey, "記錄者姓名", ir.GetValue("記錄者姓名"));
                        else
                            _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "記錄者姓名", CounselCareRecord.AuthorName, ir.GetValue("記錄者姓名"));
                        CounselCareRecord.AuthorName = ir.GetValue("記錄者姓名");
                    }
                    if (ir.Contains("記錄者"))
                    {                        

                        // 檢查記錄者如果空的用記錄者姓名比對,有比對到填入記錄者
                        if (string.IsNullOrEmpty(ir.GetValue("記錄者")))
                        {
                            if (teacherNameLoginIDDict.ContainsKey(CounselCareRecord.AuthorName))
                            {
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "記錄者", ir.GetValue("記錄者"));
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "記錄者", CounselCareRecord.AuthorID, teacherNameLoginIDDict[CounselCareRecord.AuthorName]);
                                CounselCareRecord.AuthorID = teacherNameLoginIDDict[CounselCareRecord.AuthorName];
                            }
                        }
                        else
                        {
                            if (isNew)
                                _LogTransfer.AddBatchInsertLog(StudID, insertKey, "記錄者", ir.GetValue("記錄者"));
                            else
                                _LogTransfer.AddBatchUpdateLog(StudID, CounselCareRecord.UID, "記錄者", CounselCareRecord.AuthorID, ir.GetValue("記錄者"));
                            CounselCareRecord.AuthorID = ir.GetValue("記錄者");
                        }
                    }
                    if (string.IsNullOrEmpty(CounselCareRecord.UID))
                        InsertData.Add(CounselCareRecord);
                    else
                        UpdateData.Add(CounselCareRecord);
                }
            }
            if (InsertData.Count > 0)
                _UDTTransfer.InsertCareRecordList(InsertData);

            if (UpdateData.Count > 0)
                _UDTTransfer.UpdateCareRecordList(UpdateData);

            // log
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> insertLogDict = _LogTransfer.GetBatchInsertLog();
            Dictionary<string,Dictionary<string,Dictionary<string,string>>> updateLogDict = _LogTransfer.GetBatchUpdateLog();

            string TitleName = "匯入優先關懷";
            // 處理 log 細項
            foreach (KeyValuePair<string, string> data in StudentNameDict)
            {                
                if (insertLogDict.Count > 0)
                {
                    if (insertLogDict.ContainsKey(data.Key))
                    {
                        foreach (KeyValuePair<string, Dictionary<string, string>> d1 in insertLogDict[data.Key])
                        {
                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine(TitleName+"-新增");
                            sb.AppendLine(data.Value);
                            foreach (KeyValuePair<string, string> d2 in d1.Value)
                                sb.AppendLine(d2.Key + "：" + d2.Value);

                            if (LogData.ContainsKey(data.Key))
                            {
                                LogData[data.Key].AppendLine();
                                LogData[data.Key].AppendLine(sb.ToString());
                            }
                            else
                                LogData.Add(data.Key, sb);
                            
                        }
                    }
                }

                if (updateLogDict.Count > 0)
                {
                    if (updateLogDict.ContainsKey(data.Key))
                    {
                        
                        foreach (KeyValuePair<string, Dictionary<string, string>> d1 in updateLogDict[data.Key])
                        {
                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine(TitleName+"-修改");
                            sb.AppendLine(data.Value);

                            foreach (KeyValuePair<string, string> d2 in d1.Value)
                                sb.AppendLine(d2.Key + "：" + d2.Value);

                            if (LogData.ContainsKey(data.Key))
                            {
                                LogData[data.Key].AppendLine();
                                LogData[data.Key].AppendLine(sb.ToString());
                            }
                            else
                                LogData.Add(data.Key, sb);
                        }
                    }
                }
            
            }


            // 寫入 log
            DAO.LogTransfer log = new DAO.LogTransfer();
            foreach (KeyValuePair<string, StringBuilder> data in LogData)
                log.SaveLog("輔導系統."+TitleName, "匯入", "student", data.Key, data.Value);

            StringBuilder sbT = new StringBuilder();
            sbT.AppendLine(TitleName);
            sbT.AppendLine("總共匯入" + LogData.Keys.Count + "位學生 , 共" + TotalCount + "筆");
            sbT.AppendLine("匯入學生名單..");
            foreach (KeyValuePair<string, string> data in StudentNameDict)
                if (LogData.ContainsKey(data.Key))
                    sbT.AppendLine(data.Value);

            log.SaveLog("輔導系統." + TitleName, "匯入", "student", "", sbT);
            return "";
        }

        public override void Prepare(ImportOption Option)
        {
            mOption = Option;            
            _UDTTransfer = new DAO.UDTTransfer();
        }
    }
}
