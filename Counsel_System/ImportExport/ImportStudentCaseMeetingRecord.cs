﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;
using Campus.Import;
using FISCA.UDT;
using System.Xml.Linq;

namespace Counsel_System2.ImportExport
{
    /// <summary>
    /// 匯入學生個案會議
    /// </summary>
    public class ImportStudentCaseMeetingRecord:ImportWizard
    {
        private ImportOption mOption;       
        DAO.UDTTransfer _UDTTransfer;
        public override ImportAction GetSupportActions()
        {
            return ImportAction.InsertOrUpdate;
        }

        public ImportStudentCaseMeetingRecord()
        {
            this.IsSplit = false;
            this.IsLog = false;
        }

        public override string GetValidateRule()
        {
            return Properties.Resources.ImportStudentCaseMeetingRecordVal;
        }

        public override string Import(List<IRowStream> Rows)
        {
            List<DAO.UDT_CounselCaseMeetingRecordDef> InsertData = new List<DAO.UDT_CounselCaseMeetingRecordDef>();
            List<DAO.UDT_CounselCaseMeetingRecordDef> UpdateData = new List<DAO.UDT_CounselCaseMeetingRecordDef>();
            List<DAO.UDT_CounselCaseMeetingRecordDef> HasData = new List<DAO.UDT_CounselCaseMeetingRecordDef>();

            // 取得教師帳號比對用
            Dictionary<string, string> teacherNameLoginIDDict = Utility.GetTeacherNameLoginIDStatus1();

            // -- 處理 log
            Dictionary<string, StringBuilder> LogData = new Dictionary<string, StringBuilder>();
            // 學生ID List
            List<int> studIdList = new List<int>();
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
            int TotalCount = 0, NewIdx=0;
            List<string> StudentIDList = new List<string> ();
            // 取得學生狀態對應
            foreach (IRowStream ir in Rows)
            {
                TotalCount++;
                this.ImportProgress = TotalCount;
                if (ir.Contains("學號") && ir.Contains("狀態"))
                  StudentIDList.Add(Utility.GetStudentID(ir.GetValue("學號"),ir.GetValue("狀態")).ToString ());
            }
                // 已有資料
                HasData= _UDTTransfer.GetCaseMeetingRecordListByStudentIDList(StudentIDList);


                foreach (IRowStream ir in Rows)
                { 
                    DAO.UDT_CounselCaseMeetingRecordDef CounselCaseMeetingRecord = null;
                    int sid=0;
                    if (ir.Contains("學號") && ir.Contains("狀態"))
                    {
                        string key =ir.GetValue("學號") + "_";
                        if (Global._StudentStatusDBDict.ContainsKey(ir.GetValue("狀態")))
                            sid = Utility.GetStudentID(ir.GetValue("學號"), ir.GetValue("狀態"));

                        DateTime dt;
                        // 當同一位學生有相同會議日期與會議事由，當作是更新，否則新增
                        if (DateTime.TryParse(ir.GetValue("會議日期"), out dt))
                        {
                            foreach (DAO.UDT_CounselCaseMeetingRecordDef rec in HasData.Where(x => x.StudentID == sid))
                            {
                                if (rec.MeetingDate.HasValue)
                                    if (rec.MeetingDate.Value.ToShortDateString() == dt.ToShortDateString())
                                        if (rec.MeetingCause == ir.GetValue("會議事由"))
                                            CounselCaseMeetingRecord = rec;
                                        
                                            
                            }
                        }

                        bool isNew = false;
                        if (CounselCaseMeetingRecord == null)
                        {
                            CounselCaseMeetingRecord = new DAO.UDT_CounselCaseMeetingRecordDef();
                            isNew = true;
                            NewIdx++;
                        }
                        string StudID = sid.ToString();
                        string insertKey = "A" + NewIdx;

                        // 學生編號
                        CounselCaseMeetingRecord.StudentID = sid;
                        
                        // 會議日期                          
                    if (isNew)
                        _LogTransfer.AddBatchInsertLog(StudID, insertKey, "會議日期", dt.ToShortDateString());
                    else
                        _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "會議日期", CounselCaseMeetingRecord.MeetingDate.Value.ToShortDateString(), dt.ToShortDateString());
                        CounselCaseMeetingRecord.MeetingDate = dt;

                        // 晤談老師
                        if (Global._AllTeacherNameIdDictTemp.ContainsKey(ir.GetValue("晤談老師")))
                            CounselCaseMeetingRecord.CounselTeacherID = Global._AllTeacherNameIdDictTemp[ir.GetValue("晤談老師")];

                        if (ir.Contains("個案編號"))
                        {
                            if (isNew)
                                _LogTransfer.AddBatchInsertLog(StudID, insertKey, "個案編號", ir.GetValue("個案編號"));
                            else
                                _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "個案編號", CounselCaseMeetingRecord.CaseNo, ir.GetValue("個案編號"));
                            CounselCaseMeetingRecord.CaseNo = ir.GetValue("個案編號");

                            
                        }
                        if (ir.Contains("會議時間"))
                        {
                             if (isNew)
                                _LogTransfer.AddBatchInsertLog(StudID, insertKey, "會議時間", ir.GetValue("會議時間"));
                            else
                                _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "會議時間", CounselCaseMeetingRecord.MeetigTime, ir.GetValue("會議時間"));
                            CounselCaseMeetingRecord.MeetigTime = ir.GetValue("會議時間");
                        }

                        if (ir.Contains("會議地點"))
                        {
                            if (isNew)
                                _LogTransfer.AddBatchInsertLog(StudID, insertKey, "會議地點", ir.GetValue("會議地點"));
                            else
                                _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "會議地點", CounselCaseMeetingRecord.CaseNo, ir.GetValue("會議地點"));
                            CounselCaseMeetingRecord.Place = ir.GetValue("會議地點");
                        }

                        if (ir.Contains("會議事由"))
                        {
                            if (isNew)
                                _LogTransfer.AddBatchInsertLog(StudID, insertKey, "會議事由", ir.GetValue("會議事由"));
                            else
                                _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "會議事由", CounselCaseMeetingRecord.MeetingCause, ir.GetValue("會議事由"));
                            
                            CounselCaseMeetingRecord.MeetingCause = ir.GetValue("會議事由");                            
                        }


                        StringBuilder sb1 = new StringBuilder();
                        if (ir.Contains("參與人員:學生")) if (ir.GetValue("參與人員:學生") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "學生");
                                sb1.Append(elm.ToString());

                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:學生", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:學生","","1");
                            }

                        if (ir.Contains("參與人員:家長")) if (ir.GetValue("參與人員:家長") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "家長");
                                sb1.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:家長", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:家長","","1");
                            }
                        if (ir.Contains("參與人員:專家")) if (ir.GetValue("參與人員:專家") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "專家");
                                sb1.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:專家", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:專家","","1");

                            }
                        if (ir.Contains("參與人員:醫師")) if (ir.GetValue("參與人員:醫師") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "醫師");
                                sb1.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:醫師", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:醫師", "","1");
                            }
                        if (ir.Contains("參與人員:社工人員")) if (ir.GetValue("參與人員:社工人員") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "社工人員");
                                sb1.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:社工人員", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:社工人員","","1");
                            }
                        if (ir.Contains("參與人員:導師")) if (ir.GetValue("參與人員:導師") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "導師");
                                sb1.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:導師", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:導師","","1");
                            }
                        if (ir.Contains("參與人員:教官")) if (ir.GetValue("參與人員:教官") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "教官");
                                sb1.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:教官", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:教官", "", "1");
                            }
                        if (ir.Contains("參與人員:輔導老師")) if (ir.GetValue("參與人員:輔導老師") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "輔導老師");
                                sb1.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:輔導老師", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:輔導老師", "", "1");
                            }
                        if (ir.Contains("參與人員:任課老師")) if (ir.GetValue("參與人員:任課老師") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "任課老師");
                                sb1.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:任課老師", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:任課老師", "", "1");
                            }
                        if (ir.Contains("參與人員:其它")) if (ir.GetValue("參與人員:其它") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "其它");
                                if (ir.Contains("參與人員:其它備註"))
                                    elm.SetAttributeValue("remark", ir.GetValue("參與人員:其它備註"));
                                sb1.Append(elm.ToString());

                                if (isNew)
                                {
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:其它", "1");
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "參與人員:其它備註", ir.GetValue("參與人員:其它備註"));
                                }
                                else
                                {
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:其它", "", "1");
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "參與人員:其它", "", ir.GetValue("參與人員:其它備註"));
                                }
                            }

                        CounselCaseMeetingRecord.Attendees = sb1.ToString();

                        StringBuilder sb2 = new StringBuilder();

                        if (ir.Contains("輔導方式:暫時結案")) if (ir.GetValue("輔導方式:暫時結案") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "暫時結案");
                                sb2.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導方式:暫時結案", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導方式:暫時結案", "", "1");
                            }
                        if (ir.Contains("輔導方式:專案輔導")) if (ir.GetValue("輔導方式:專案輔導") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "專案輔導");
                                sb2.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導方式:專案輔導", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導方式:專案輔導", "", "1");

                            }
                        if (ir.Contains("輔導方式:導師輔導")) if (ir.GetValue("輔導方式:導師輔導") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "導師輔導");
                                sb2.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導方式:導師輔導", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導方式:導師輔導", "", "1");

                            }
                        if (ir.Contains("輔導方式:轉介")) if (ir.GetValue("輔導方式:轉介") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "轉介");

                                if (ir.Contains("輔導方式:轉介備註"))
                                    elm.SetAttributeValue("remark", ir.GetValue("輔導方式:轉介備註"));
                                sb2.Append(elm.ToString());
                                if (isNew)
                                {
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導方式:轉介", "1");
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導方式:轉介備註", ir.GetValue("輔導方式:轉介備註"));
                                }
                                else
                                {
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導方式:轉介備註", "", ir.GetValue("輔導方式:轉介備註"));
                                }
                            }


                        if (ir.Contains("輔導方式:就醫")) if (ir.GetValue("輔導方式:就醫") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "就醫");

                                if (ir.Contains("輔導方式:就醫備註"))
                                    elm.SetAttributeValue("remark", ir.GetValue("輔導方式:就醫備註"));
                                sb2.Append(elm.ToString());

                                if (isNew)
                                {
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導方式:就醫", "1");
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導方式:就醫備註", "1");
                                }
                                else
                                {
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導方式:就醫", "", "1");
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導方式:就醫備註", "", ir.GetValue("輔導方式:就醫備註"));
                                }
                            }


                        if (ir.Contains("輔導方式:其它")) if (ir.GetValue("輔導方式:其它") == "1")
                        {
                            XElement elm = new XElement("Item");
                            elm.SetAttributeValue("name", "其它");


                            if (ir.Contains("輔導方式:其它備註"))
                                elm.SetAttributeValue("remark", ir.GetValue("輔導方式:其它備註"));
                            sb2.Append(elm.ToString());
                            
                            if (isNew)
                            {
                                _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導方式:其它", "1");
                                _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導方式:其它備註", "1");
                            }
                            else
                            {
                                _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導方式:其它", "", "1");
                                _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導方式:其它備註", "", ir.GetValue("輔導方式:其它備註"));
                            }
                        }

                        CounselCaseMeetingRecord.CounselType = sb2.ToString();

                        StringBuilder sb3 = new StringBuilder();
                        
                        if (ir.Contains("輔導歸類:違規")) if (ir.GetValue("輔導歸類:違規") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "違規");
                                sb3.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:違規", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:違規", "", "1");

                            }
                        if (ir.Contains("輔導歸類:遲曠")) if (ir.GetValue("輔導歸類:遲曠") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "遲曠");
                                sb3.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:遲曠", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:遲曠", "", "1");

                            }
                        if (ir.Contains("輔導歸類:學習")) if (ir.GetValue("輔導歸類:學習") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "學習");
                                sb3.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:學習", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:學習", "", "1");
                            }
                        if (ir.Contains("輔導歸類:生涯")) if (ir.GetValue("輔導歸類:生涯") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "生涯");
                                sb3.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:生涯", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:生涯", "", "1");
                            }
                        if (ir.Contains("輔導歸類:人")) if (ir.GetValue("輔導歸類:人") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "人");
                                sb3.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:人", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:人", "", "1");
                            }
                        if (ir.Contains("輔導歸類:休退轉")) if (ir.GetValue("輔導歸類:休退轉") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "休退轉");
                                sb3.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:休退轉", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:休退轉", "", "1");

                            }
                        if (ir.Contains("輔導歸類:家庭")) if (ir.GetValue("輔導歸類:家庭") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "家庭");
                                sb3.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:家庭", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:家庭", "", "1");
                            }
                        if (ir.Contains("輔導歸類:師生")) if (ir.GetValue("輔導歸類:師生") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "師生");
                                sb3.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:師生", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:師生", "", "1");
                            }
                        if (ir.Contains("輔導歸類:情感")) if (ir.GetValue("輔導歸類:情感") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "情感");
                                sb3.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:情感", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:情感", "", "1");
                            }
                        if (ir.Contains("輔導歸類:精神")) if (ir.GetValue("輔導歸類:精神") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "精神");
                                sb3.Append(elm.ToString());
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:精神", "1");
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:精神", "", "1");
                            }

                        if (ir.Contains("輔導歸類:其它")) if (ir.GetValue("輔導歸類:其它") == "1")
                            {
                                XElement elm = new XElement("Item");
                                elm.SetAttributeValue("name", "其它");
                                if (ir.Contains("輔導歸類:其它備註"))
                                    elm.SetAttributeValue("remark", ir.GetValue("輔導歸類:其它備註"));
                                sb3.Append(elm.ToString());
                                if (isNew)
                                {
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:其它", "1");
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "輔導歸類:其它備註", ir.GetValue("輔導歸類:其它備註"));
                                }
                                else
                                {
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:其它", "", "1");
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "輔導歸類:其它備註", "", ir.GetValue("輔導歸類:其它備註"));
                                }
                            }

                        CounselCaseMeetingRecord.CounselTypeKind = sb3.ToString();

                        if(ir.Contains("內容要點"))
                        {                            
                            if (isNew)
                                _LogTransfer.AddBatchInsertLog(StudID, insertKey, "內容要點", ir.GetValue("內容要點"));
                            else
                                _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "內容要點", CounselCaseMeetingRecord.ContentDigest, ir.GetValue("內容要點"));
                            CounselCaseMeetingRecord.ContentDigest=ir.GetValue("內容要點");
                        }
                    if (ir.Contains("記錄者姓名"))
                    {
                        if (isNew)
                            _LogTransfer.AddBatchInsertLog(StudID, insertKey, "記錄者姓名", ir.GetValue("記錄者姓名"));
                        else
                            _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "記錄者姓名", CounselCaseMeetingRecord.AuthorName, ir.GetValue("記錄者姓名"));
                        CounselCaseMeetingRecord.AuthorName = ir.GetValue("記錄者姓名");
                    }
                    if (ir.Contains("記錄者"))
                    {

                        // 檢查記錄者如果空的用記錄者姓名比對,有比對到填入記錄者
                        if (string.IsNullOrEmpty(ir.GetValue("記錄者")))
                        {
                            if (teacherNameLoginIDDict.ContainsKey(CounselCaseMeetingRecord.AuthorName))
                            {
                                if (isNew)
                                    _LogTransfer.AddBatchInsertLog(StudID, insertKey, "記錄者", ir.GetValue("記錄者"));
                                else
                                    _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "記錄者", CounselCaseMeetingRecord.AuthorID, teacherNameLoginIDDict[CounselCaseMeetingRecord.AuthorName]);
                                CounselCaseMeetingRecord.AuthorID = teacherNameLoginIDDict[CounselCaseMeetingRecord.AuthorName];
                            }
                        }
                        else
                        {
                            if (isNew)
                                _LogTransfer.AddBatchInsertLog(StudID, insertKey, "記錄者", ir.GetValue("記錄者"));
                            else
                                _LogTransfer.AddBatchUpdateLog(StudID, CounselCaseMeetingRecord.UID, "記錄者", CounselCaseMeetingRecord.AuthorID, ir.GetValue("記錄者"));
                            CounselCaseMeetingRecord.AuthorID = ir.GetValue("記錄者");
                        }
                    }

                        if (string.IsNullOrEmpty(CounselCaseMeetingRecord.UID))
                            InsertData.Add(CounselCaseMeetingRecord);
                        else
                            UpdateData.Add(CounselCaseMeetingRecord);
                    }
                }
                if (InsertData.Count > 0)
                    _UDTTransfer.InsertCaseMeetingRecordList(InsertData);

                if (UpdateData.Count > 0)
                    _UDTTransfer.UpdateCaseMeetingRecordList(UpdateData);

                // log
                Dictionary<string, Dictionary<string, Dictionary<string, string>>> insertLogDict = _LogTransfer.GetBatchInsertLog();
                Dictionary<string, Dictionary<string, Dictionary<string, string>>> updateLogDict = _LogTransfer.GetBatchUpdateLog();

                string TitleName = "匯入個案會議";
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
                                sb.AppendLine(TitleName + "-新增");
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
                                sb.AppendLine(TitleName + "-修改");
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
                    log.SaveLog("輔導系統." + TitleName, "匯入", "student", data.Key, data.Value);

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
