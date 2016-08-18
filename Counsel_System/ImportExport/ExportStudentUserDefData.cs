using System.Collections.Generic;
using SmartSchool.API.PlugIn;

namespace Counsel_System2.ImportExport
{
    // 匯出自訂資料欄位
    public class ExportStudentUserDefData : SmartSchool.API.PlugIn.Export.Exporter 
    {
        List<string> ExportItemList;
        DAO.UDTTransfer _UDTTransfer;

        public ExportStudentUserDefData()
        {
            this.Image = null;
            _UDTTransfer = new DAO.UDTTransfer();
            this.Text = "匯出輔導自訂資料欄位";
            ExportItemList = new List<string>();
            ExportItemList.Add("欄位名稱");
            ExportItemList.Add("值");
            ExportItemList.Add("狀態");
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange(ExportItemList);
            wizard.ExportPackage += delegate(object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                int RowCount = 0;

                List<DAO.UDT_CounselUserDefDataDef> CounselUserDefDataList = _UDTTransfer.GetCounselUserDefineDataByStudentIDList(e.List);

                foreach (DAO.UDT_CounselUserDefDataDef udd in CounselUserDefDataList)
                {
                    RowData row = new RowData();
                    row.ID = udd.StudentID.ToString ();

                    foreach (string field in e.ExportFields)
                    {
                        if (wizard.ExportableFields.Contains(field))
                        {
                            switch (field)
                            {
                                case "欄位名稱": row.Add(field, udd.FieldName); break;
                                case "值": row.Add(field, udd.Value); break;
                                case "狀態":
                                    row.Add(field, udd.StudentStatus);
                                    break;

                            }
                        }
                    
                    }
                    RowCount++;
                    e.Items.Add(row);                
                }
            };
        }
    }
}
