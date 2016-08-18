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
    public partial class AddABCardSubjectNameForm : FISCA.Presentation.Controls.BaseForm
    {
        List<string> _SubjectNameList;
        string _SubjectName = "";

        public AddABCardSubjectNameForm(List<string> SubjectNameList)
        {
            InitializeComponent();
            this.MaximumSize = this.MinimumSize = this.Size;
            _SubjectNameList = SubjectNameList;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 取得新增的綜合表現名稱
        /// </summary>
        /// <returns></returns>
        public string GetAddSubjectName()
        {
            return _SubjectName;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            _SubjectName = txtSubjectName.Text;

            // 檢查名稱是否重複
            if (_SubjectNameList.Contains(_SubjectName))
            {
                FISCA.Presentation.Controls.MsgBox.Show("系統內已有相同名稱，無法新增.");
                return;
            }

            this.DialogResult = System.Windows.Forms.DialogResult.Yes;
        }
    }
}
