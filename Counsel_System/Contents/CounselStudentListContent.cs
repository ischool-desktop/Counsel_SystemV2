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
    /// <summary>
    /// 學生.輔導學生
    /// </summary>
    [FISCA.Permission.FeatureCode(PermissionCode.認輔老師及輔導老師_資料項目, "認輔/輔導老師")]
    public partial class CounselStudentListContent : FISCA.Presentation.DetailContent
    {
        DAO.UDTTransfer _UDTTransfer;
        List<DAO.UDT_CounselStudent_ListDef> _CounselStudent_List;
        BackgroundWorker _bgWorker;
        Dictionary<int, DAO.CounselTeacherRecord> _CounselTeacherDict;
        bool isBGBusy = false;
        // 排序用
        List<ListViewItem> _lviList;

        public CounselStudentListContent()
        {
            InitializeComponent();
            _lviList = new List<ListViewItem>();
            Group = "認輔/輔導老師";
            _UDTTransfer = new DAO.UDTTransfer();
            _CounselTeacherDict = new Dictionary<int, DAO.CounselTeacherRecord>();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWork_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWork_RunWorkerCompleted);
            EventHub.CounselChanged += new EventHandler(EventHub_CounselChanged);
        }

        void EventHub_CounselChanged(object sender, EventArgs e)
        {
            _BGRun();
        }

        private void _BGRun()
        {
            if (_bgWorker.IsBusy)
                isBGBusy = true;
            else
                _bgWorker.RunWorkerAsync();
        }

        void _bgWork_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (isBGBusy)
            {
                isBGBusy = false;
                _bgWorker.RunWorkerAsync();
                return;
            }
            LoadDataToListView();
        }

        void _bgWork_DoWork(object sender, DoWorkEventArgs e)
        {
            _CounselStudent_List = _UDTTransfer.GetCounselStudentListID(PrimaryKey);
            _CounselTeacherDict = Utility.GetAllCounselTeacherDict();            
        }

        private void LoadDataToListView()
        {          

            this.Loading = false;
            lvStudentList.Items.Clear();
            _lviList.Clear();
            ListViewItem lvi = null;
            int countA = 0, countB = 0;
            foreach (DAO.UDT_CounselStudent_ListDef rec in _CounselStudent_List)
            {                
                if (_CounselTeacherDict.ContainsKey(rec.TeacherTagID))
                {
                    lvi = new ListViewItem();
                    lvi.Tag = rec;

                    if (_CounselTeacherDict[rec.TeacherTagID].counselTeacherType == DAO.CounselTeacherRecord.CounselTeacherType.認輔老師)
                        countA++;
                    if (_CounselTeacherDict[rec.TeacherTagID].counselTeacherType == DAO.CounselTeacherRecord.CounselTeacherType.輔導老師)
                        countB++;
                    
                    lvi.Text = _CounselTeacherDict[rec.TeacherTagID].TeacherFullName;                    
                    lvi.SubItems.Add(_CounselTeacherDict[rec.TeacherTagID].counselTeacherType.ToString());
                                     
                    _lviList.Add(lvi);
                }
            }

            // sort by TeacherName ascending
            _lviList = (from data in _lviList orderby data.Text ascending select data).ToList();

            lvStudentList.Items.AddRange(_lviList.ToArray());

            lblTeacherCount.Text = "認輔教師:" + countA + "人  輔導教師:" + countB + "人";
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            this.Loading = true;

            _BGRun();
        }

        /// <summary>
        /// 認輔教師
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetCounselTeacherA_Click(object sender, EventArgs e)
        {
            List<string> StudentIDList = new List<string>();
            StudentIDList.Add(PrimaryKey);
            Forms.SetCounselTeacherForm stf = new Forms.SetCounselTeacherForm(DAO.CounselTeacherRecord.CounselTeacherType.認輔老師, StudentIDList);
            stf.ShowDialog();
            _BGRun();
        }

        /// <summary>
        /// 輔導老師
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetCounselTeacherB_Click(object sender, EventArgs e)
        {
            List<string> StudentIDList = new List<string>();
            StudentIDList.Add(PrimaryKey);
            Forms.SetCounselTeacherForm stf = new Forms.SetCounselTeacherForm(DAO.CounselTeacherRecord.CounselTeacherType.輔導老師, StudentIDList);
            stf.ShowDialog();
            _BGRun();
        }

        private void btnRemoveCounselTeacher_Click(object sender, EventArgs e)
        {
            List<DAO.UDT_CounselStudent_ListDef> DelList = new List<DAO.UDT_CounselStudent_ListDef>();

            if (lvStudentList.SelectedItems.Count > 0)
            {

                foreach (ListViewItem lvi in lvStudentList.SelectedItems)
                {
                    DAO.UDT_CounselStudent_ListDef data = lvi.Tag as DAO.UDT_CounselStudent_ListDef;
                    if (data == null)
                        continue;

                    DelList.Add(data);
                }
                if (DelList.Count > 0)
                {
                    if (FISCA.Presentation.Controls.MsgBox.Show("確定移除" + DelList.Count + "位教師?", "移除教師", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        _UDTTransfer.DeleteCounselStudentList(DelList);
                        _BGRun();
                    }
                }
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show("請選擇老師.");

        }
    }
}
