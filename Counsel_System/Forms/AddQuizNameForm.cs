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
    public partial class AddQuizNameForm : FISCA.Presentation.Controls.BaseForm
    {
        List<string> _QuizNameList;
        string _AddName = "";
        public AddQuizNameForm(List<string> QuizNameList)
        {
            InitializeComponent();
            _QuizNameList = new List<string>();
            this.MaximumSize = this.MinimumSize = this.Size;
            _QuizNameList = QuizNameList;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (_QuizNameList.Contains(txtQuizName.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("系統內已有相同名稱，無法新增.");
                return;
            }
            else
            {
                // 檢查空白
                if (string.IsNullOrWhiteSpace(txtQuizName.Text))
                {
                    FISCA.Presentation.Controls.MsgBox.Show("名稱不能空白，請輸入名稱.");
                    return;
                }

                _AddName = txtQuizName.Text;
            }
            this.DialogResult = System.Windows.Forms.DialogResult.Yes;
        }

        /// <summary>
        /// 取得新增測驗名稱
        /// </summary>
        /// <returns></returns>
        public string GetAddQuizName()
        {
            return _AddName.Trim();
        }
    }
}
