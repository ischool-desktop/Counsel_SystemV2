using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using K12.Data;

namespace Counsel_System2.Contents
{
    /// <summary>
    /// 教師.認輔教師
    /// </summary>
    [FISCA.Permission.FeatureCode(PermissionCode.認輔學生_資料項目, "認輔學生")]
    public partial class CounselStudentListAContent : FISCA.Presentation.DetailContent
    {
        int? _TeacherEntityID;
        BackgroundWorker _bgWorker;
        Dictionary<int, StudentRecord> _StudDict = new Dictionary<int, StudentRecord>();
        List<DAO.UDT_CounselStudent_ListDef> _CounselStudent;
        DAO.UDTTransfer _UDTTransfer;
        bool isBGBusy = false;
        List<ListViewItem> _lviList;
        
        public CounselStudentListAContent()
        {
            InitializeComponent();
            _lviList = new List<ListViewItem>();
            Group = "認輔學生";
            _CounselStudent = new List<DAO.UDT_CounselStudent_ListDef>();
            _UDTTransfer = new DAO.UDTTransfer();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;

            EventHub.CounselChanged += new EventHandler(EventHub_CounselChanged);
        }

        void EventHub_CounselChanged(object sender, EventArgs e)
        {
            _BGRun();
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

        private void LoadDataToListView()
        {
            this.Loading = false;
            // 判斷是輔導老師才能使用
            if (_TeacherEntityID.HasValue)
            {
                btnAddTempStudent.Enabled = true;
                btnRemoveStudent.Enabled = true;
            }
            else
            {
                btnAddTempStudent.Enabled = false;
                btnRemoveStudent.Enabled = false;
            }

            lvStudentList.Items.Clear();
            _lviList.Clear();
            ListViewItem lvi = null;
            foreach (DAO.UDT_CounselStudent_ListDef rec in _CounselStudent)
            {
                if(_StudDict.ContainsKey(rec.StudentID))
                {
                    
                    lvi = new ListViewItem();
                    StudentRecord studrec=_StudDict[rec.StudentID];
                    lvi.Tag =rec;
                    
                    if(studrec.Class != null )
                        lvi.Text =studrec.Class.Name ;
                    else
                        lvi.Text ="";
                    if(studrec.SeatNo.HasValue )
                        lvi.SubItems.Add(studrec.SeatNo.Value.ToString ());
                    else
                        lvi.SubItems.Add("");

                    lvi.SubItems.Add(studrec.Name );
                    lvi.SubItems.Add(studrec.StudentNumber);
                    lvi.SubItems.Add(studrec.Status.ToString ());
                    _lviList.Add(lvi);
                    //lvStudentList.Items.Add(lvi);
                }            
            }
            
            // sort
            int num;
            _lviList = (from data in _lviList orderby data.Text ascending, int.TryParse(data.SubItems[colSeatNo.Index].Text,out num) ascending,data.SubItems[colName.Index].Text ascending select data).ToList();
            lvStudentList.Items.AddRange(_lviList.ToArray());

            lblStudentCount.Text = "學生人數：" + lvStudentList.Items.Count;
            
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            //_TeacherEntityID = null;
            _CounselStudent.Clear();
            _TeacherEntityID = _UDTTransfer.GetTeacherTagIDByTeacherID(PrimaryKey, DAO.CounselTeacherRecord.CounselTeacherType.認輔老師.ToString ());
            if(_TeacherEntityID.HasValue )
            _CounselStudent = _UDTTransfer.GetCounselStudentListByTeacherTagID(_TeacherEntityID.Value);


            //foreach (DAO.CounselTeacherRecord ctr in Utility.GetCounselTeacherDictByTeacherID(PrimaryKey))
            //{
            //    if (ctr.counselTeacherType == DAO.CounselTeacherRecord.CounselTeacherType.認輔老師)
            //        _TeacherEntityID = ctr.TeacherTag_ID;
            //}

            //if (_TeacherEntityID.HasValue)
            //{
            //    _CounselStudent = _UDTTransfer.GetCounselStudentListByTeacherTagID(_TeacherEntityID.Value);
                _StudDict.Clear();
                List<string> sidList = (from data in _CounselStudent select data.StudentID.ToString()).ToList();
                foreach (StudentRecord stud in Student.SelectByIDs(sidList))
                {
                    int sid = int.Parse(stud.ID);
                    if (!_StudDict.ContainsKey(sid))
                        _StudDict.Add(sid, stud);                
                }
            //}
        }

        private void btnRemoveStudent_Click(object sender, EventArgs e)
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
                    if (FISCA.Presentation.Controls.MsgBox.Show("確定移除" + DelList.Count + "位學生?", "移除所選學生", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        _UDTTransfer.DeleteCounselStudentList(DelList);
                        _bgWorker.RunWorkerAsync();
                    }
                }
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show("請選擇資料.");

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

        private void btnAddTempStudent_PopupOpen(object sender, EventArgs e)
        {
            CreateStudentMenuItem();
        }

        private void CreateStudentMenuItem()
        {
            btnAddTempStudent.SubItems.Clear();
            if (K12.Presentation.NLDPanels.Student.TempSource.Count == 0)
            {
                LabelItem item = new LabelItem("No", "沒有任何學生在待處理");
                btnAddTempStudent.SubItems.Add(item);
                return;
            }

            List<StudentRecord> studTempRecList = Student.SelectByIDs(K12.Presentation.NLDPanels.Student.TempSource);
            
            foreach (StudentRecord stud in studTempRecList)
            {
                string strclassname = "", strSeatNo = ""; ;
                if(stud.Class !=null )
                {
                    strclassname=stud.Class.Name;
                }
                if (stud.SeatNo.HasValue)
                    strSeatNo = stud.SeatNo.Value.ToString();
                ButtonItem item = new ButtonItem (stud.ID,stud.Name+"("+strclassname+")");
                item.Tooltip = "班級：" + strclassname + "\n座號：" + strSeatNo + "\n學號：" + stud.StudentNumber;
                item.Tag = stud;                
                item.Click += new EventHandler(item_Click);
                btnAddTempStudent.SubItems.Add(item);
            }
        
        }

        void item_Click(object sender, EventArgs e)
        {
            StudentRecord stud=null;
            ButtonItem bt = sender as ButtonItem;
            if(bt !=null )
                stud = bt.Tag as StudentRecord;
            if (stud != null && _TeacherEntityID.HasValue)
            {
                // 檢查是否加入
                int sid = int.Parse(stud.ID);                
                if (!_StudDict.ContainsKey(sid))
                {

                    List<DAO.UDT_CounselStudent_ListDef> dataList = new List<DAO.UDT_CounselStudent_ListDef>();
                    DAO.UDT_CounselStudent_ListDef data = new DAO.UDT_CounselStudent_ListDef();
                    data.StudentID = sid;
                    data.TeacherTagID = _TeacherEntityID.Value;
                    dataList.Add(data);
                    _UDTTransfer.InsertCounselStudentList(dataList);
                    _BGRun();
                }
            }

        }

        private void btnAddTempStudent_Click(object sender, EventArgs e)
        {
            CreateStudentMenuItem();

            List<DAO.UDT_CounselStudent_ListDef> dataList = new List<DAO.UDT_CounselStudent_ListDef>();

            foreach (object obj in btnAddTempStudent.SubItems)
            {
                StudentRecord stud = null;
                ButtonItem bt = obj as ButtonItem;
                if (bt != null)
                    stud = bt.Tag as StudentRecord;

                if (stud != null && _TeacherEntityID.HasValue)
                {
                    // 當已經加入跳過
                    int sid = int.Parse(stud.ID);
                    if (_StudDict.ContainsKey(sid))
                        continue;

                    DAO.UDT_CounselStudent_ListDef data = new DAO.UDT_CounselStudent_ListDef();
                    data.StudentID = sid;
                    data.TeacherTagID = _TeacherEntityID.Value;
                    dataList.Add(data);
                }
            }
            _UDTTransfer.InsertCounselStudentList(dataList);
            _BGRun();
        }

        private void CounselStudentListAContent_Load(object sender, EventArgs e)
        {

        }
    }
}
