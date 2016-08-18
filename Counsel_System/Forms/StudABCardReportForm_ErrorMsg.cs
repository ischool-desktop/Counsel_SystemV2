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
    public partial class StudABCardReportForm_ErrorMsg : FISCA.Presentation.Controls.BaseForm
    {
        StringBuilder _sb = new StringBuilder();
        public StudABCardReportForm_ErrorMsg(StringBuilder sb)
        {
            InitializeComponent();
            _sb = sb;
        }

        private void StudABCardReportForm_ErrorMsg_Load(object sender, EventArgs e)
        {
            txtMsg.Text = _sb.ToString();
        }
    }
}
