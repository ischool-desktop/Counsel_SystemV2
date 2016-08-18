using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace Counsel_System2.Contents
{
    // [FISCA.Permission.FeatureCode(PermissionCode.輔導晤談紀錄_資料項目, "晤談紀錄")]
    [FISCA.Permission.FeatureCode(PermissionCode.輔導晤談紀錄_資料項目, "輔導紀錄")]
    public partial class StudInterviewDataContent : FISCA.Presentation.DetailContent
    {
         List<DAO.UDT_CounselStudentInterviewRecordDef> _StudentInterviewRecordList;
         DAO.UDTTransfer _UDTTransfer;
         BackgroundWorker _bgWorker;
         Dictionary<int, string> _TeacherIDNameDict;
         bool isBGBusy = false;

        public StudInterviewDataContent()
        {
            InitializeComponent();
            Group = "輔導紀錄";
            _StudentInterviewRecordList = new List<DAO.UDT_CounselStudentInterviewRecordDef>();
            _UDTTransfer = new DAO.UDTTransfer();
            _TeacherIDNameDict = new Dictionary<int, string>();
            
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

        private void StudInterviewDataContent_Load(object sender, EventArgs e)
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
         
            LoadDataToDataGridView();
        }

        private void LoadDataToDataGridView()
        {
            this.Loading = false;
            lvInterview.Items.Clear();
            ListViewItem lvi = null;
            foreach (DAO.UDT_CounselStudentInterviewRecordDef rec in _StudentInterviewRecordList)
            {
                lvi = new ListViewItem();
                lvi.Tag = rec;

                // 晤談日期
                if (rec.InterviewDate.HasValue)
                    lvi.Text = rec.InterviewDate.Value.ToShortDateString ();
                else
                    lvi.Text = " ";
               
                // 晤談對象
                lvi.SubItems.Add(rec.IntervieweeType);

                // 老師
                if (_TeacherIDNameDict.ContainsKey(rec.TeacherID))
                    lvi.SubItems.Add(_TeacherIDNameDict[rec.TeacherID]);
                else
                    lvi.SubItems.Add("");

                // 事由
                lvi.SubItems.Add(rec.Cause);                              

                lvInterview.Items.Add(lvi);
            }
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _StudentInterviewRecordList = _UDTTransfer.GetCounselStudentInterviewRecordByStudentID(PrimaryKey).OrderByDescending(x=>x.InterviewDate).ToList();
            _TeacherIDNameDict = Utility.GetTeacherIDNameDict();

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            DAO.UDT_CounselStudentInterviewRecordDef StudentInterviewRecord = new DAO.UDT_CounselStudentInterviewRecordDef();
            StudentInterviewRecord.StudentID = int.Parse(PrimaryKey);
            Forms.StudInterviewDataForm sidf = new Forms.StudInterviewDataForm(StudentInterviewRecord,Forms.StudInterviewDataForm.AccessType.Insert);
            sidf.ShowDialog();
            _BGRun();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (lvInterview.SelectedItems.Count == 1)
            {
                DAO.UDT_CounselStudentInterviewRecordDef rec = lvInterview.SelectedItems[0].Tag as DAO.UDT_CounselStudentInterviewRecordDef;
                if (rec != null)
                {
                    Forms.StudInterviewDataForm sidf = new Forms.StudInterviewDataForm(rec,Forms.StudInterviewDataForm.AccessType.Update);
                    sidf.ShowDialog();
                }
                _BGRun();
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show("請選擇資料.");
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (lvInterview.SelectedItems.Count > 0)
            {   
                DAO.UDT_CounselStudentInterviewRecordDef studInterviewRec = lvInterview.SelectedItems[0].Tag as DAO.UDT_CounselStudentInterviewRecordDef;
                if (studInterviewRec != null)
                {
                    DAO.LogTransfer logTransfer = new DAO.LogTransfer();
                    string teacherName = lvInterview.SelectedItems[0].SubItems[colTeacherID.Index].Text;
                    K12.Data.StudentRecord studRec=K12.Data.Student.SelectByID(PrimaryKey);                    
                    StringBuilder logData = new StringBuilder();
                    logData.AppendLine("刪除" + Utility.ConvertString1(studRec));
                    // 取得 XML 解析後
                    Dictionary<string, string> item_AttendessDict = Utility.GetConvertCounselXMLVal_Attendees(studInterviewRec.Attendees);
                    Dictionary<string, string> item_CounselTypeDict = Utility.GetConvertCounselXMLVal_CounselType(studInterviewRec.CounselType);
                    Dictionary<string, string> item_CounselTypeKindDict = Utility.GetConvertCounselXMLVal_CounselTypeKind(studInterviewRec.CounselTypeKind);

                    
                    if (FISCA.Presentation.Controls.MsgBox.Show("請問是否確定是刪除晤談紀錄?", "刪除晤談紀錄", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        // log                        
                        logData.AppendLine("晤談老師："+ teacherName);
                        logData.AppendLine("晤談對象："+ studInterviewRec.IntervieweeType);
                        logData.AppendLine("晤談方式："+ studInterviewRec.InterviewType);
                        if (studInterviewRec.InterviewDate.HasValue)
                            logData.AppendLine("日期："+ studInterviewRec.InterviewDate.Value.ToShortDateString());
                        logData.AppendLine("時間："+ studInterviewRec.InterviewTime);
                        logData.AppendLine("地點："+ studInterviewRec.Place);
                        logData.AppendLine("晤談編號："+ studInterviewRec.InterviewNo);                        
                        logData.AppendLine("晤談事由："+ studInterviewRec.Cause);

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

                        logData.AppendLine("內容要點："+ studInterviewRec.ContentDigest);
                        logData.AppendLine("記錄者："+ studInterviewRec.AuthorID);
                        logData.AppendLine("記錄者姓名："+ studInterviewRec.AuthorName);

                        _UDTTransfer.DeleteCounselStudentInterviewRecord(studInterviewRec);

                        logTransfer.SaveLog("學生.輔導晤談紀錄-刪除", "刪除", "student", PrimaryKey, logData);

                        
                        _BGRun();
                    }
                }
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show ("請選擇資料.");
        }


    }
}
