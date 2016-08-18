using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.API.PlugIn;

namespace Counsel_System2.ImportExport
{
    // 匯出個案會議
    class ExportCaseMeetingRecord:SmartSchool.API.PlugIn.Export.Exporter
    {
        // 可勾選選項
        List<string> ExportItemList;
        List<string> _XmlItemList;
        DAO.UDTTransfer _UDTTransfer;
        public ExportCaseMeetingRecord()
        {
            this.Image = null;
            this.Text = "匯出個案會議";
            ExportItemList = new List<string>();
            _XmlItemList = Utility.GetCounselXmlItemList_CaseMeeting();
            _UDTTransfer = new DAO.UDTTransfer();
            ExportItemList.Add("個案編號");
            ExportItemList.Add("晤談老師");
            ExportItemList.Add("會議日期");
            ExportItemList.Add("會議時間");
            ExportItemList.Add("會議事由");
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
               // 透過學生編號取得個案會議
               List<DAO.UDT_CounselCaseMeetingRecordDef> CounselCaseMeetingRecordList = _UDTTransfer.GetCaseMeetingRecordListByStudentIDList(e.List);

               // 取得教師ID與名稱對照
               Dictionary<int, string> TeacherNameDict = Utility.GetTeacherIDNameDict();
               Dictionary<string, string> item_AttendessDict = new Dictionary<string, string>();
               Dictionary<string, string> item_CounselTypeDict = new Dictionary<string, string>();
               Dictionary<string, string> item_CounselTypeKindDict = new Dictionary<string, string>();

               // 學生ID List
               List<int> studIdList = (from data in CounselCaseMeetingRecordList select data.StudentID).ToList();

               // 取得學生名稱log 用
               Dictionary<string, string> StudentNameDict = new Dictionary<string, string>();
               foreach (KeyValuePair<int, string> data in Utility.GetConvertStringDict1fromDB(studIdList))
                   StudentNameDict.Add(data.Key.ToString(), data.Value);

               // 記錄匯出 log 細項用
               Dictionary<string, StringBuilder> logData = new Dictionary<string, StringBuilder>();

               // 匯出總筆數
               int totalCount = 0;


               foreach (DAO.UDT_CounselCaseMeetingRecordDef ccmrd in CounselCaseMeetingRecordList)
               {
                   RowData row = new RowData();
                   row.ID = ccmrd.StudentID.ToString();

                   // 取得 XML 解析後
                   item_AttendessDict = Utility.GetConvertCounselXMLVal_Attendees(ccmrd.Attendees);
                   item_CounselTypeDict = Utility.GetConvertCounselXMLVal_CounselType(ccmrd.CounselType);
                   item_CounselTypeKindDict = Utility.GetConvertCounselXMLVal_CounselTypeKind(ccmrd.CounselTypeKind);



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
                               case "個案編號":
                                   row.Add(field, ccmrd.CaseNo);
                                   break;
                               case "晤談老師":
                                   // 需轉換
                                   if(TeacherNameDict.ContainsKey(ccmrd.CounselTeacherID))
                                       row.Add(field, TeacherNameDict[ccmrd.CounselTeacherID]);
                                   break;
                               case "會議日期":
                                   if(ccmrd.MeetingDate.HasValue )
                                    row.Add(field, ccmrd.MeetingDate.Value.ToShortDateString());
                                   break;
                               case "會議時間":
                                   row.Add(field, ccmrd.MeetigTime);
                                   break;
                               case "會議事由":
                                   row.Add(field, ccmrd.MeetingCause);
                                   break;             
                                   
                               case "內容要點": 
                                   row.Add(field,ccmrd.ContentDigest);
                                   break;

                               case "記錄者":
                                   row.Add(field, ccmrd.AuthorID);
                                   break;

                               case "記錄者姓名":
                                   row.Add(field, ccmrd.AuthorName);
                                   break;

                               case "狀態":
                                   row.Add(field, ccmrd.StudentStatus);
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
