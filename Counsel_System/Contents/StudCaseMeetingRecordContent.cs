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
    [FISCA.Permission.FeatureCode(PermissionCode.輔導個案會議_資料項目, "個案會議")]
    public partial class StudCaseMeetingRecordContent : FISCA.Presentation.DetailContent
    {
        List<DAO.UDT_CounselCaseMeetingRecordDef> _StudentCaseMeetingRecordList;
        DAO.UDTTransfer _UDTTransfer;
        BackgroundWorker _bgWorker;
        bool isBGBusy = false;

        public StudCaseMeetingRecordContent()
        {
            InitializeComponent();
            Group = "個案會議";
            _StudentCaseMeetingRecordList = new List<DAO.UDT_CounselCaseMeetingRecordDef>();
            _UDTTransfer = new DAO.UDTTransfer();
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


        private void LoadDataToListView()
        {
            this.Loading = false;
            lvCaseMeeting.Items.Clear();
            ListViewItem lvi = null;
            foreach (DAO.UDT_CounselCaseMeetingRecordDef rec in _StudentCaseMeetingRecordList)
            {
                lvi = new ListViewItem();
                lvi.Tag = rec;
                if (rec.MeetingDate.HasValue)
                    lvi.Text = rec.MeetingDate.Value.ToString();
                else
                    lvi.Text = "";

                lvi.SubItems.Add(rec.AuthorID);
                lvi.SubItems.Add(rec.MeetingCause);
                lvCaseMeeting.Items.Add(lvi);
            }
        }

        private void StudCaseMeetingRecordContent_Load(object sender, EventArgs e)
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
            _StudentCaseMeetingRecordList = _UDTTransfer.GetCaseMeetingRecordListByStudentID(PrimaryKey).OrderByDescending(x=>x.MeetingDate).ToList ();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (lvCaseMeeting.SelectedItems.Count == 1)
            {
                
                DAO.UDT_CounselCaseMeetingRecordDef CaseMeetingRecord = lvCaseMeeting.SelectedItems[0].Tag as DAO.UDT_CounselCaseMeetingRecordDef;
                if (CaseMeetingRecord != null)
                {
                    DAO.LogTransfer logTransfer = new DAO.LogTransfer();
                    
                    K12.Data.StudentRecord studRec = K12.Data.Student.SelectByID(PrimaryKey);
                    StringBuilder logData = new StringBuilder();
                    logData.AppendLine("刪除" + Utility.ConvertString1(studRec));

                    // 取得 XML 解析後
                    Dictionary<string, string> item_AttendessDict = Utility.GetConvertCounselXMLVal_Attendees(CaseMeetingRecord.Attendees);
                    Dictionary<string, string> item_CounselTypeDict = Utility.GetConvertCounselXMLVal_CounselType(CaseMeetingRecord.CounselType);
                    Dictionary<string, string> item_CounselTypeKindDict = Utility.GetConvertCounselXMLVal_CounselTypeKind(CaseMeetingRecord.CounselTypeKind);

                    if (FISCA.Presentation.Controls.MsgBox.Show("請問是否確定刪除個案會議?", "刪除個案會議", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        // log
                        logData.AppendLine("個案編號：" + CaseMeetingRecord.CaseNo);
                    
                        if (CaseMeetingRecord.MeetingDate.HasValue)
                            logData.AppendLine("會議日期"+ CaseMeetingRecord.MeetingDate.Value.ToShortDateString());

                        K12.Data.TeacherRecord tRec = K12.Data.Teacher.SelectByID(CaseMeetingRecord.CounselTeacherID.ToString());
                        if (tRec != null)
                            if(string.IsNullOrEmpty(tRec.Nickname))
                                logData.AppendLine("晤談老師：" + tRec.Name);
                            else
                                logData.AppendLine("晤談老師：" + tRec.Name+"("+tRec.Nickname+")");

                        logData.AppendLine("會議時間：" + CaseMeetingRecord.MeetigTime);
                        logData.AppendLine("會議事由：" + CaseMeetingRecord.MeetingCause);
                        logData.AppendLine("會議地點：" + CaseMeetingRecord.Place);
                        logData.AppendLine("內容要點：" + CaseMeetingRecord.ContentDigest);
                        logData.AppendLine("記錄者：" + CaseMeetingRecord.AuthorID);
                        logData.AppendLine("記錄者姓名：" + CaseMeetingRecord.AuthorName);

                        logData.AppendLine("參與人員：");
                        foreach (KeyValuePair<string, string> data in item_AttendessDict)
                            if (!string.IsNullOrEmpty(data.Value))
                                logData.AppendLine(data.Key + ":" + data.Value);

                        logData.AppendLine("輔導方式：");
                        foreach (KeyValuePair<string, string> data in item_CounselTypeDict)
                            if (!string.IsNullOrEmpty(data.Value))
                                logData.AppendLine(data.Key + ":" + data.Value);

                        logData.AppendLine("輔導歸類：");
                        foreach (KeyValuePair<string, string> data in item_CounselTypeKindDict)
                            if (!string.IsNullOrEmpty(data.Value))
                                logData.AppendLine(data.Key + ":" + data.Value);


                        _UDTTransfer.DeleteCaseMeetingRecord(CaseMeetingRecord);
                        logTransfer.SaveLog("學生.輔導個案會議-刪除", "刪除", "student", PrimaryKey, logData);
                        _BGRun();
                    }
                }
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show("請選擇資料.");

            
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (lvCaseMeeting.SelectedItems.Count == 1)
            {
                DAO.UDT_CounselCaseMeetingRecordDef CaseMeetingRecord = lvCaseMeeting.SelectedItems[0].Tag as DAO.UDT_CounselCaseMeetingRecordDef;
                if (CaseMeetingRecord != null)
                {
                    Forms.StudCaseMeetingRecordForm scmrf = new Forms.StudCaseMeetingRecordForm(CaseMeetingRecord, Forms.StudCaseMeetingRecordForm.accessType.Edit);
                    if(scmrf.ShowDialog()== DialogResult.OK)
                        _BGRun();
                }
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show("請選擇資料.");
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            DAO.UDT_CounselCaseMeetingRecordDef CaseMeetingRecord = new DAO.UDT_CounselCaseMeetingRecordDef();
            CaseMeetingRecord.StudentID = int.Parse(PrimaryKey);
            Forms.StudCaseMeetingRecordForm scmrf = new Forms.StudCaseMeetingRecordForm(CaseMeetingRecord, Forms.StudCaseMeetingRecordForm.accessType.Insert);
            if(scmrf.ShowDialog()== DialogResult.OK)
                _BGRun();
        }
    }
}
