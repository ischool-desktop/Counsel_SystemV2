using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.Import;
using Campus.DocumentValidator;
using FISCA.UDT;
using System.Xml.Linq;

namespace Counsel_System2.ImportExport
{
    /// <summary>
    /// 匯入輔導使用者自訂欄位
    /// </summary>
    public class ImportStudentUserDefData : ImportWizard
    {
        private ImportOption mOption;
        DAO.UDTTransfer _UDTTransfer;

        public ImportStudentUserDefData()
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
            return Properties.Resources.ImportStudentUserDefDataVal;
        }

        public override string Import(List<IRowStream> Rows)
        {
            List<DAO.UDT_CounselUserDefDataDef> InsertData = new List<DAO.UDT_CounselUserDefDataDef>();
            List<DAO.UDT_CounselUserDefDataDef> UpdateData = new List<DAO.UDT_CounselUserDefDataDef>();
            List<DAO.UDT_CounselUserDefDataDef> HasData = new List<DAO.UDT_CounselUserDefDataDef>();

            List<string> StudentIDList = new List<string>();
            // 取得學生狀態對應
            foreach (IRowStream ir in Rows)
            {
                if (ir.Contains("學號") && ir.Contains("狀態"))
                    StudentIDList.Add(Utility.GetStudentID(ir.GetValue("學號"), ir.GetValue("狀態")).ToString());
            }
            // 已有資料
            HasData = _UDTTransfer.GetCounselUserDefineDataByStudentIDList(StudentIDList);


            foreach (IRowStream ir in Rows)
            {
                DAO.UDT_CounselUserDefDataDef CounselUserDefineData = null;
                int sid = 0;
                if (ir.Contains("學號") && ir.Contains("狀態"))
                {
                    string key = ir.GetValue("學號") + "_";
                    if (Global._StudentStatusDBDict.ContainsKey(ir.GetValue("狀態")))
                        sid = Utility.GetStudentID(ir.GetValue("學號"), ir.GetValue("狀態"));

                        foreach (DAO.UDT_CounselUserDefDataDef rec in HasData.Where(x => x.StudentID == sid))
                        {
                            if (rec.FieldName == ir.GetValue("欄位名稱"))
                                        CounselUserDefineData = rec;
                        }

                    if (CounselUserDefineData == null)
                        CounselUserDefineData = new DAO.UDT_CounselUserDefDataDef();

                    // 學生編號
                    CounselUserDefineData.StudentID = sid;
                    // 欄位名稱
                    CounselUserDefineData.FieldName = ir.GetValue("欄位名稱");
                    // 值
                    CounselUserDefineData.Value = ir.GetValue("值");


                    if (string.IsNullOrEmpty(CounselUserDefineData.UID))
                        InsertData.Add(CounselUserDefineData);
                    else
                        UpdateData.Add(CounselUserDefineData);
                }
            }
            if (InsertData.Count > 0)
                _UDTTransfer.InsertCounselUsereDefinfDataList(InsertData);

            if (UpdateData.Count > 0)
                _UDTTransfer.UpdateCounselUserDefineDataList(UpdateData);

            return "";
        }

        public override void Prepare(ImportOption Option)
        {
            mOption = Option;
            _UDTTransfer = new DAO.UDTTransfer();
        }
    }
}
