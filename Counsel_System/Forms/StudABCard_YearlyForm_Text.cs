using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Counsel_System2.DAO;

namespace Counsel_System2.Forms
{
    public partial class StudABCard_YearlyForm_Text : FISCA.Presentation.Controls.BaseForm
    {
        UDTYearlyDataDef _data;
        public StudABCard_YearlyForm_Text(UDTYearlyDataDef data)
        {
            InitializeComponent();
            _data = data;            
        }

        private void StudABCard_YearlyForm_Text_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            LoadData();
        }

        /// <summary>
        /// 取得資料
        /// </summary>
        /// <returns></returns>
        public UDTYearlyDataDef GetData()
        {
            SetData();
            return _data;
        }

        private void LoadData()
        {
            txt01.Text = _data.G1;
            txt02.Text = _data.G2;
            txt03.Text = _data.G3;
            txt04.Text = _data.G4;
            txt05.Text = _data.G5;
            txt06.Text = _data.G6;            
        }

        public void SetTitleName(string name)
        {
            this.Text = name;
        }

        private void SetData()
        {
            _data.G1 = txt01.Text;
            _data.G2 = txt02.Text;
            _data.G3 = txt03.Text;
            _data.G4 = txt04.Text;
            _data.G5 = txt05.Text;
            _data.G6 = txt06.Text;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            SetData();
            this.Close();
        }
    }
}
