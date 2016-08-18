using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.API.PlugIn;

namespace Counsel_System2.ImportExport
{
    // 匯出晤談紀錄
    class ExportInterViewData:SmartSchool.API.PlugIn.Export.Exporter
    {
        // 可勾選選項
        List<string> ExportItemList;
        DAO.UDTTransfer _UDTTransfer;
        List<string> _XmlItemList;
        public ExportInterViewData()
        {
            this.Image = null;
            this.Text = "匯出晤談紀錄";
            ExportItemList = new List<string>();
            _XmlItemList = Utility.GetCounselXmlItemList_InterView();
            _UDTTransfer = new DAO.UDTTransfer();
            ExportItemList.Add("晤談編號");
            ExportItemList.Add("晤談老師");
            ExportItemList.Add("晤談方式");
            ExportItemList.Add("晤談對象");
            ExportItemList.Add("日期");
            ExportItemList.Add("時間");
            ExportItemList.Add("地點");
            ExportItemList.Add("晤談事由");
            ExportItemList.AddRange(_XmlItemList.ToArray());
            ExportItemList.Add("內容要點");
            ExportItemList.Add("記錄者");
            ExportItemList.Add("記錄者姓名");
            ExportItemList.Add("狀態");
        
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange(ExportItemList);
            wizard.ExportPackage += delegate(object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
           {
               // 依學生ID取得晤談紀錄
                List<DAO.UDT_CounselStudentInterviewRecordDef> _StudentInterviewRecordList = _UDTTransfer.GetCounselStudentInterviewRecordByStudentIDList(e.List);

                // 取得教師ID與名稱對照
                Dictionary<int, string> TeacherNameDict = Utility.GetTeacherIDNameDict();

               Dictionary<string, string> item_AttendessDict = new Dictionary<string, string>();
                Dictionary<string, string> item_CounselTypeDict = new Dictionary<string, string>();
                Dictionary<string, string> item_CounselTypeKindDict = new Dictionary<string, string>();

                // 學生ID List
                List<int> studIdList = (from data in _StudentInterviewRecordList select data.StudentID).ToList();

                // 取得學生名稱log 用
                Dictionary<string, string> StudentNameDict = new Dictionary<string, string>();
                foreach (KeyValuePair<int, string> data in Utility.GetConvertStringDict1fromDB(studIdList))
                    StudentNameDict.Add(data.Key.ToString(), data.Value);

                // 記錄匯出 log 細項用
                Dictionary<string, StringBuilder> logData = new Dictionary<string, StringBuilder>();

                // 匯出總筆數
                int totalCount = 0;

                foreach (DAO.UDT_CounselStudentInterviewRecordDef csir in _StudentInterviewRecordList)
                {

                    // 取得 XML 解析後
                    item_AttendessDict = Utility.GetConvertCounselXMLVal_Attendees(csir.Attendees);
                    item_CounselTypeDict = Utility.GetConvertCounselXMLVal_CounselType(csir.CounselType);
                    item_CounselTypeKindDict = Utility.GetConvertCounselXMLVal_CounselTypeKind(csir.CounselTypeKind);

                    RowData row = new RowData();
                    row.ID = csir.StudentID.ToString();
                     foreach (string field in e.ExportFields)
                     {
                         if (wizard.ExportableFields.Contains(field))
                         {
                             // 參與人員
                             if (item_AttendessDict.ContainsKey(field))
                                 row.Add(field, item_AttendessDict[field]);

                             // 輔導方式
                             if (item_CounselTypeDict.ContainsKey(field))
                                 row.Add(field, item_CounselTypeDict[field]);

                             // 輔導歸類
                             if (item_CounselTypeKindDict.ContainsKey(field))
                                 row.Add(field, item_CounselTypeKindDict[field]);

                             switch (field)
                             {
                                 case "晤談編號":
                                     row.Add(field, csir.InterviewNo);
                                     break;
                                 case "晤談老師":
                                     // 需轉換
                                   if(TeacherNameDict.ContainsKey(csir.TeacherID))
                                       row.Add(field, TeacherNameDict[csir.TeacherID]);
                                     break;
                                 case "晤談方式":
                                     row.Add(field, csir.InterviewType);
                                     break;
                                 case "晤談對象":
                                     row.Add(field, csir.IntervieweeType);
                                     break;
                                 case "日期":
                                     if(csir.InterviewDate.HasValue )
                                        row.Add(field, csir.InterviewDate.Value.ToShortDateString());
                                     break;
                                 case "時間":
                                     row.Add(field, csir.InterviewTime);
                                     break;
                                 case "地點":
                                     row.Add(field, csir.Place);
                                     break;
                                 case "晤談事由":
                                     row.Add(field, csir.Cause);
                                     break;
                                 case "參與人員": 
                                     // 解析 XML
                                     row.Add(field,Utility.AttendeesXMLToString(csir.Attendees));
                                     break;
                                 case "輔導方式":
                                     // 解析 XML
                                     row.Add(field, Utility.CounselTypeXMLToString(csir.CounselType));
                                     break;
                                 case "輔導歸類":
                                     // 解析 XML
                                     row.Add(field, Utility.CounselTypeKindXMLToString(csir.CounselTypeKind));
                                     break;
                                 case "內容要點":
                                     row.Add(field, csir.ContentDigest);
                                     break;

                                 case "記錄者":
                                     row.Add(field, csir.AuthorID);
                                     break;

                                 case "記錄者姓名":
                                     row.Add(field, csir.AuthorName);
                                     break;

                                 case "狀態":
                                     row.Add(field, csir.StudentStatus);
                                     break;

                             }
                         }
                     }
                     e.Items.Add(row);
                }
                // 處理 log 細項
                foreach (RowData rd in e.Items)
                {
                    // 收集 log
                    StringBuilder sb = new StringBuilder();
                    if (StudentNameDict.ContainsKey(rd.ID))
                        sb.AppendLine(StudentNameDict[rd.ID]);
                    sb.AppendLine(this.Text);
                    foreach (KeyValuePair<string, string> data in rd)
                        sb.AppendLine(data.Key + "：" + data.Value);

                    // 加入 log
                    if (logData.ContainsKey(rd.ID))
                    {
                        sb.AppendLine();
                        logData[rd.ID].AppendLine(sb.ToString());
                    }
                    else
                        logData.Add(rd.ID, sb);
                }

                // 寫入 log
                DAO.LogTransfer log = new DAO.LogTransfer();
                foreach (KeyValuePair<string, StringBuilder> data in logData)
                    log.SaveLog("輔導系統." + this.Text, "匯出", "student", data.Key, data.Value);

                StringBuilder sbT = new StringBuilder();
                sbT.AppendLine(this.Text);
                sbT.AppendLine("總共匯出" + logData.Keys.Count + "位學生 , 共" + totalCount + "筆");
                sbT.AppendLine("匯出學生名單..");
                foreach (KeyValuePair<string, string> data in StudentNameDict)
                    if (logData.ContainsKey(data.Key))
                        sbT.AppendLine(data.Value);

                log.SaveLog("輔導系統." + this.Text, "匯出", "student", "", sbT);
           };
        }
    }
}
