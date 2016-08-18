using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Counsel_System2.Forms
{
    public partial class SetCounselTeacherForm : FISCA.Presentation.Controls.BaseForm 
    {
        DAO.CounselTeacherRecord.CounselTeacherType _CounselTeacherType;
        List<DAO.UDT_CounselStudent_ListDef> _hasCounselStudent_List;
        List<int> _StudentIDList;
        DAO.UDTTransfer _UDTTransfer;

        public SetCounselTeacherForm(DAO.CounselTeacherRecord.CounselTeacherType Type,List<string> StudentIDList)
        {
            InitializeComponent();
            _UDTTransfer = new DAO.UDTTransfer ();
            this.MaximumSize = this.MinimumSize = this.Size;
            _CounselTeacherType = Type;
            _StudentIDList = new List<int>();
            _hasCounselStudent_List = new List<DAO.UDT_CounselStudent_ListDef>();
            _StudentIDList = (from data in StudentIDList select int.Parse(data)).ToList();
            if(StudentIDList.Count>0)
                _hasCounselStudent_List = _UDTTransfer.GetCounselStudentListIDs(StudentIDList);

            this.Text = "指定" + Type.ToString();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            List<DAO.UDT_CounselStudent_ListDef> InsertData = new List<DAO.UDT_CounselStudent_ListDef>();
            foreach (ListViewItem lvi in lvTeacherName.CheckedItems)
            {
                DAO.CounselTeacherRecord ctr = lvi.Tag as DAO.CounselTeacherRecord;
                if (ctr == null)
                    continue;

                foreach (int id in _StudentIDList)
                {   
                    // 檢查是否已經指定過
                    int hasDataCount = (from data in _hasCounselStudent_List where data.StudentID == id && data.TeacherTagID == ctr.TeacherTag_ID select data).ToList().Count;
                    // 沒有設定
                    if (hasDataCount == 0)
                    {
                        DAO.UDT_CounselStudent_ListDef data = new DAO.UDT_CounselStudent_ListDef();
                        data.StudentID = id;
                        data.TeacherTagID = ctr.TeacherTag_ID;
                        InsertData.Add(data);                    
                    }
                }
            }
            if (InsertData.Count > 0)
                _UDTTransfer.InsertCounselStudentList(InsertData);
            if(lvTeacherName.CheckedItems.Count>0)
                FISCA.Presentation.Controls.MsgBox.Show("指定完成.");            

            EventHub.OnCounselChanged();
            Close();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        public void LoadDataToListView()
        {
            lvTeacherName.Items.Clear();
            ListViewItem lvi = null;
            foreach (DAO.CounselTeacherRecord rec in Utility.GetCounselTeacherByType(_CounselTeacherType))
            {
                lvi = new ListViewItem();
                lvi.Tag = rec;
                lvi.Text = rec.TeacherFullName;
                lvTeacherName.Items.Add(lvi);
            }
        
        }

        private void SetCounselTeacherForm_Load(object sender, EventArgs e)
        {
            LoadDataToListView();
        }

    }
}
