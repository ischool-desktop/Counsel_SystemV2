using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Counsel_System2.Contents
{
    [FISCA.Permission.FeatureCode(PermissionCode.輔導優先關懷紀錄_資料項目, "優先關懷")]
    public partial class StudCareRecordContent : FISCA.Presentation.DetailContent
    {
        
        DAO.UDTTransfer _UDTTransfer;
        BackgroundWorker _bgWorker;
        List<DAO.UDT_CounselCareRecordDef> _CareRecordList;
        bool isBGBusy = false;
        public StudCareRecordContent()
        {
            InitializeComponent();
            Group = "優先關懷";
            _UDTTransfer = new DAO.UDTTransfer();
            _CareRecordList = new List<DAO.UDT_CounselCareRecordDef>();            
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            this.Loading = true;

            _BGRun();
        }

        private void _BGRun()
        {
            if (_bgWorker.IsBusy)
                isBGBusy = true;
            else
                _bgWorker.RunWorkerAsync();
        }

        private void StudCareRecordContent_Load(object sender, EventArgs e)
        {
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (isBGBusy)
            {
                isBGBusy = false;
                _bgWorker.RunWorkerAsync();
                return;
            }           
             LoadDataToListView();
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _CareRecordList = _UDTTransfer.GetCareRecordsByStudentID(PrimaryKey).OrderByDescending(x=>x.FileDate).ToList ();
            
        }

        private void LoadDataToListView()
        {
            this.Loading = false;
            lvCareRecord.Items.Clear();
            ListViewItem lvi = null;
            foreach (DAO.UDT_CounselCareRecordDef rec in _CareRecordList)
            {
                lvi = new ListViewItem();
                lvi.Tag = rec;
                lvi.Text = rec.AuthorID;

                if (rec.FileDate.HasValue)
                    lvi.SubItems.Add(rec.FileDate.Value.ToShortDateString());
                else
                    lvi.SubItems.Add("");
                lvi.SubItems.Add(rec.CaseCategory);
                lvCareRecord.Items.Add(lvi);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (lvCareRecord.SelectedItems.Count == 1)
            {
                

                DAO.UDT_CounselCareRecordDef careRec = lvCareRecord.SelectedItems[0].Tag as DAO.UDT_CounselCareRecordDef;
                if (careRec != null)
                {
                    DAO.LogTransfer logTransfer = new DAO.LogTransfer();
                    K12.Data.StudentRecord studRec = K12.Data.Student.SelectByID(PrimaryKey);
                    StringBuilder logData = new StringBuilder();
                    logData.AppendLine("刪除" + Utility.ConvertString1(studRec));

                    if (FISCA.Presentation.Controls.MsgBox.Show("請問是否確定是刪除優先關懷紀錄?", "刪除優先關懷紀錄", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        // log
                        logData.AppendFormat("代號：" + careRec.CodeName);                        
                        if (careRec.FileDate.HasValue)
                            logData.AppendLine("立案日期："+careRec.FileDate.Value.ToShortDateString());

                        logData.AppendLine("個案類別：" + careRec.CaseCategory);
                        logData.AppendLine("個案類別備註：" + careRec.CaseCategoryRemark);
                        logData.AppendLine("個案來源：" + careRec.CaseOrigin);
                        logData.AppendLine("個案來源備註：" + careRec.CaseOriginRemark);
                        logData.AppendLine("優勢能力及財力：" + careRec.Superiority);
                        logData.AppendLine("弱勢能力及阻力：" + careRec.Weakness);
                        logData.AppendLine("輔導人員輔導目標：" + careRec.CounselGoal);
                        logData.AppendLine("校外協輔機構：" + careRec.OtherInstitute);
                        logData.AppendLine("輔導人員輔導方式：" + careRec.CounselType);
                        logData.AppendLine("協同輔導人員協助導師事項：" + careRec.AssistedMatter);
                        logData.AppendLine("記錄者：" + careRec.AuthorID);
                        logData.AppendLine("記錄者姓名：" + careRec.AuthorName);
                        
                        _UDTTransfer.DeleteCareRecord(careRec);

                        logTransfer.SaveLog("學生.輔導優先關懷-刪除", "刪除", "student", PrimaryKey, logData);

                        _BGRun();
                    }
                }
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show("請選擇資料.");
            
          
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            DAO.UDT_CounselCareRecordDef CareRecord = new DAO.UDT_CounselCareRecordDef();
            CareRecord.StudentID = int.Parse(PrimaryKey);
            Forms.StudCareRecordForm scrf = new Forms.StudCareRecordForm(CareRecord,Forms.StudCareRecordForm.accessType.Insert);
            if(scrf.ShowDialog() == DialogResult.OK)
                _BGRun();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (lvCareRecord.SelectedItems.Count == 1)
            { 
                DAO.UDT_CounselCareRecordDef careRec = lvCareRecord.SelectedItems[0].Tag as DAO.UDT_CounselCareRecordDef;
                if (careRec != null)
                {
                    Forms.StudCareRecordForm scrf = new Forms.StudCareRecordForm(careRec, Forms.StudCareRecordForm.accessType.Edit);
                    if(scrf.ShowDialog()== DialogResult.OK)
                        _BGRun();
                }
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show("請選擇資料.");
        }
    }
}
