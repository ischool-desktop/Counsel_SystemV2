using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.API.PlugIn;

namespace Counsel_System2.ImportExport
{
    // 匯出優先關懷
    class ExportCareRecord:SmartSchool.API.PlugIn.Export.Exporter 
    {
        // 可勾選選項
        List<string> ExportItemList;
        DAO.UDTTransfer _UDTTransfer;
        public ExportCareRecord()
        {
            this.Image = null;
            this.Text = "匯出優先關懷";
            ExportItemList = new List<string>();
            _UDTTransfer = new DAO.UDTTransfer();
            ExportItemList.Add("代號");
            ExportItemList.Add("立案日期");
            ExportItemList.Add("個案類別");
            ExportItemList.Add("個案類別備註");
            ExportItemList.Add("個案來源");
            ExportItemList.Add("個案來源備註");
            ExportItemList.Add("優勢能力及財力");
            ExportItemList.Add("弱勢能力及財力");
            ExportItemList.Add("輔導人員輔導目標");
            ExportItemList.Add("校外協輔機構");
            ExportItemList.Add("輔導人員輔導方式");
            ExportItemList.Add("協同輔導人員協助導師事項");
            ExportItemList.Add("記錄者");
            ExportItemList.Add("記錄者姓名");
            ExportItemList.Add("狀態");
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange(ExportItemList);
            wizard.ExportPackage += delegate(object sender,SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                // 依學生ID取得優先關懷
                List<DAO.UDT_CounselCareRecordDef> _CounselCareRecordList = _UDTTransfer.GetCareRecordsByStudentIDList(e.List);
                
                // 學生ID List
                List<int> studIdList = (from data in _CounselCareRecordList select data.StudentID).ToList();

                // 取得學生名稱log 用
                Dictionary<string, string> StudentNameDict = new Dictionary<string, string>();
                foreach (KeyValuePair<int, string> data in Utility.GetConvertStringDict1fromDB(studIdList))
                    StudentNameDict.Add(data.Key.ToString(), data.Value);

                // 記錄匯出 log 細項用
                Dictionary<string, StringBuilder> logData = new Dictionary<string, StringBuilder>();

                // 匯出總筆數
                int totalCount = 0;

                foreach (DAO.UDT_CounselCareRecordDef ccrd in _CounselCareRecordList)
                {
                    RowData row = new RowData();
                    totalCount++;

                    row.ID = ccrd.StudentID.ToString();
                    foreach (string field in e.ExportFields)
                    {
                        if (wizard.ExportableFields.Contains(field))
                        { 
                            switch(field)
                            {
                                case "代號":
                                    row.Add(field, ccrd.CodeName);
                                    break;
                                case "立案日期":
                                    if(ccrd.FileDate.HasValue)
                                        row.Add(field, ccrd.FileDate.Value.ToShortDateString());
                                    break;
                                case "個案類別":
                                    row.Add(field, ccrd.CaseCategory);
                                    break;
                                case "個案類別備註":
                                    row.Add(field, ccrd.CaseCategoryRemark);
                                    break;
                                case "個案來源":
                                    row.Add(field, ccrd.CaseOrigin);
                                    break;
                                case "個案來源備註":
                                    row.Add(field, ccrd.CaseOriginRemark);
                                    break;
                                case "優勢能力及財力":
                                    row.Add(field, ccrd.Superiority);
                                    break;
                                case "弱勢能力及財力":
                                    row.Add(field, ccrd.Weakness);
                                    break;
                                case "輔導人員輔導目標":
                                    row.Add(field, ccrd.CounselGoal);
                                    break;
                                case "校外協輔機構":
                                    row.Add(field, ccrd.OtherInstitute);
                                    break;
                                case "輔導人員輔導方式":
                                    row.Add(field, ccrd.CounselType);                                    
                                    break;
                                case "協同輔導人員協助導師事項":
                                    row.Add(field, ccrd.AssistedMatter);
                                    break;
                                case "記錄者":
                                    row.Add(field, ccrd.AuthorID);
                                    break;
                                case "記錄者姓名":
                                    row.Add(field, ccrd.AuthorName);
                                    break;
                                case "狀態":
                                    row.Add(field, ccrd.StudentStatus);
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

                StringBuilder sbT= new StringBuilder ();
                sbT.AppendLine(this.Text);
                sbT.AppendLine("總共匯出"+logData.Keys.Count+"位學生 , 共"+totalCount+"筆");
                sbT.AppendLine("匯出學生名單..");
                foreach (KeyValuePair<string, string> data in StudentNameDict)
                    if (logData.ContainsKey(data.Key))
                        sbT.AppendLine(data.Value);

                log.SaveLog("輔導系統." + this.Text, "匯出","student","",sbT);

            };
        }
    }
}
