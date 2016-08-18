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
     [FISCA.Permission.FeatureCode(PermissionCode.輔導相關測驗_資料項目, "心理測驗")]
    public partial class StudQuizDataContent : FISCA.Presentation.DetailContent
    {
        private BackgroundWorker _worker;
        DAO.UDTTransfer _UDTTransfer;
        List<DAO.UDT_StudQuizDataDef> _StudQuizDataList;
        bool isBGBusy = false;
        List<ListViewItem> _lviList;
        DAO.LogTransfer _LogTransfer;
        public StudQuizDataContent()
        {
            InitializeComponent();
            _lviList = new List<ListViewItem>();
            _UDTTransfer = new DAO.UDTTransfer();
            _StudQuizDataList = new List<DAO.UDT_StudQuizDataDef>();
            _LogTransfer = new DAO.LogTransfer();
            Group = "心理測驗";
        }

        private void _BGRun()
        {
            if (_worker.IsBusy)
                isBGBusy = true;
            else
                _worker.RunWorkerAsync();
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            this.Loading = true;
            _BGRun();
        }

        private void StudQuizDataContent_Load(object sender, EventArgs e)
        {
            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(_worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_worker_RunWorkerCompleted);
        }

        void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (isBGBusy)
            {
                isBGBusy = false;
                _worker.RunWorkerAsync();
                return;
            }
                LoadDataToListView();

        }

        void _worker_DoWork(object sender, DoWorkEventArgs e)
        {
            _StudQuizDataList = _UDTTransfer.GetStudQuizDataByStudentID(PrimaryKey);
        }

        private void LoadDataToListView()
        {
            this.Loading = false;
            lvStudQuizData.Items.Clear();
            ListViewItem lvi = null;
            _lviList.Clear();                    
            foreach (DAO.UDT_StudQuizDataDef sqd in _StudQuizDataList)
            {
                lvi = new ListViewItem();
                lvi.Tag = sqd;
                DAO.UDT_QuizDef qd = _UDTTransfer.GetQuizDataByID(sqd.QuizID.ToString ());
                lvi.Text = qd.QuizName;
                if (sqd.ImplementationDate.HasValue)
                    lvi.SubItems.Add(sqd.ImplementationDate.Value.ToShortDateString());
                else
                    lvi.SubItems.Add("");
                if (sqd.AnalysisDate.HasValue)
                    lvi.SubItems.Add(sqd.AnalysisDate.Value.ToShortDateString());
                else
                    lvi.SubItems.Add("");

                //lvStudQuizData.Items.Add(lvi);
                _lviList.Add(lvi);
            }

            // sort by ImplementationDate descending
            _lviList = (from data in _lviList orderby data.SubItems[colImplementationDate.Index].Text descending select data).ToList();
            lvStudQuizData.Items.AddRange(_lviList.ToArray());
        }


        private void LoadSubFormData()
        {
            if (lvStudQuizData.SelectedItems.Count == 1)
            {
                DAO.UDT_StudQuizDataDef sqd = lvStudQuizData.SelectedItems[0].Tag as DAO.UDT_StudQuizDataDef;
                if (sqd != null)
                {
                    Forms.StudQuizDataForm sqdf = new Forms.StudQuizDataForm(sqd,Forms.StudQuizDataForm.EditMode.Edit,PrimaryKey);
                    sqdf.ShowDialog();
                }
                _BGRun();
            }
        }

        private void lvStudQuizData_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            LoadSubFormData();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            LoadSubFormData();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            // 刪除
            if (lvStudQuizData.SelectedItems.Count == 1)
            {
                DAO.UDT_StudQuizDataDef sqd = lvStudQuizData.SelectedItems[0].Tag as DAO.UDT_StudQuizDataDef;
                if (sqd != null)
                {
                    
                    if (FISCA.Presentation.Controls.MsgBox.Show("請問刪除測驗資料","刪除測驗資料", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    { 
                        List<DAO.UDT_StudQuizDataDef> delData = new List<DAO.UDT_StudQuizDataDef> ();
                        delData.Add(sqd);
                        K12.Data.StudentRecord rec = K12.Data.Student.SelectByID(PrimaryKey);
                        StringBuilder logData = new StringBuilder();
                        logData.AppendLine("刪除" + Utility.ConvertString1(rec));
                        logData.AppendLine("測驗名稱「 " + lvStudQuizData.SelectedItems[0].SubItems[0].Text + " 」");
                        if (sqd.ImplementationDate.HasValue)
                            logData.AppendLine("實施日期：" + sqd.ImplementationDate.Value.ToShortDateString());
                        if (sqd.AnalysisDate.HasValue)
                            logData.AppendLine("解析日期：" + sqd.AnalysisDate.Value.ToShortDateString());
                        XElement elmContent = Utility.ConvertStringToXelm1(sqd.Content);
                            if (elmContent != null)
                            {
                                foreach (XElement elm in elmContent.Elements("Item"))
                                {                                    
                                    if (elm.Attribute("name") != null && elm.Attribute("value") != null)
                                        logData.AppendLine("項目名稱：" + elm.Attribute("name").Value + " , 測驗結果：" + elm.Attribute("value").Value);
                                }
                            }

                        _UDTTransfer.DeleteStudQuizDataLlist(delData);                       
                        
                        // log
                        _LogTransfer.SaveLog("學生.輔導相關測驗-刪除", "刪除", "student", PrimaryKey, logData);

                        _BGRun();
                    }
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // 新增
            Forms.StudQuizDataForm sqdf = new Forms.StudQuizDataForm(null, Forms.StudQuizDataForm.EditMode.Insert,PrimaryKey);
            sqdf.ShowDialog();
            _BGRun();
        }
    }
}
