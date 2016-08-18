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
    public partial class DelStudentQuizDataForm : FISCA.Presentation.Controls.BaseForm
    {
        BackgroundWorker _bgWork;
        List<DAO.UDT_QuizDef> _AllQuizData;
        Dictionary<string, int> _QuizIDDict;
        List<DAO.UDT_StudQuizDataDef> _StudentQuizDataList;
        DAO.UDTTransfer _UDTTransfer;
        List<string> _SelectStudentIDList;

        /// <summary>
        /// 刪除學生測驗資料
        /// </summary>
        public DelStudentQuizDataForm(List<string> StudentIDList)
        {
            InitializeComponent();
            _SelectStudentIDList = StudentIDList;
            _QuizIDDict = new Dictionary<string, int>();
            _StudentQuizDataList = new List<DAO.UDT_StudQuizDataDef>();
            _AllQuizData = new List<DAO.UDT_QuizDef>();
            _UDTTransfer = new DAO.UDTTransfer();
            this.MaximumSize = this.MinimumSize = this.Size;
            this.btnDel.Enabled = false;
            _bgWork = new BackgroundWorker();
            _bgWork.DoWork += new DoWorkEventHandler(_bgWork_DoWork);
            _bgWork.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWork_RunWorkerCompleted);
        }

        void _bgWork_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.btnDel.Enabled = true;
            // 控制項放入資料
            cbxQuizName.Items.AddRange(_QuizIDDict.Keys.ToArray());
            this.lblMsg.Visible = false;
        }

        void _bgWork_DoWork(object sender, DoWorkEventArgs e)
        {
            // 取得測驗資料
            _AllQuizData = _UDTTransfer.GetAllQuizData();
            // 取得所選學生測驗資料
            _StudentQuizDataList = _UDTTransfer.GetStudQuizDataByStudentIDList(_SelectStudentIDList);
            _QuizIDDict.Clear();
            // 比對測驗資料名稱

            foreach(DAO.UDT_QuizDef QN in _AllQuizData)
            {
                int count = 0;
                int id=int.Parse(QN.UID);
                foreach (DAO.UDT_StudQuizDataDef StudQ in _StudentQuizDataList.Where(x=>x.QuizID==id))
                {
                    count++;
                }
            
                // 有資料
                if (count > 0)
                {
                    if (!_QuizIDDict.ContainsKey(QN.QuizName))
                        _QuizIDDict.Add(QN.QuizName, id);
                }
            }
        }

        private void DelStudentQuizDataForm_Load(object sender, EventArgs e)
        {
            lblMsg.Visible = true;
            _bgWork.RunWorkerAsync();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            btnDel.Enabled = false;
            if (string.IsNullOrEmpty(cbxQuizName.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇測驗.");
                btnDel.Enabled = true;
                return;
            }

            // 取得所選測驗ID
            int QID = 0;
            if (_QuizIDDict.ContainsKey(cbxQuizName.Text))
                QID = _QuizIDDict[cbxQuizName.Text];

            List<DAO.UDT_StudQuizDataDef> DelList = new List<DAO.UDT_StudQuizDataDef>();
            DelList = (from data in _StudentQuizDataList where data.QuizID == QID select data).ToList();

            if (DelList.Count > 0)
            {
                if (FISCA.Presentation.Controls.MsgBox.Show("請問確定刪除學生測驗資料?", "刪除學生測驗資料", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                {
                    _UDTTransfer.DeleteStudQuizDataLlist(DelList);
                    FISCA.Presentation.Controls.MsgBox.Show("刪除完成.");
                    this.Close();
                }
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("沒有測驗資料.");
            }
            btnDel.Enabled = true;
        }
    }
}
