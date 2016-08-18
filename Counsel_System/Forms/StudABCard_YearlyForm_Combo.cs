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
    public partial class StudABCard_YearlyForm_Combo : FISCA.Presentation.Controls.BaseForm
    {
        UDTYearlyDataDef _data;

        public StudABCard_YearlyForm_Combo(UDTYearlyDataDef data)
        {
            InitializeComponent();
            _data = data;
        }

        /// <summary>
        /// 設標題
        /// </summary>
        /// <param name="name"></param>
        public void SetFormTitle(string name)
        {
            this.Text = name;
        }
        
        /// <summary>
        /// 設一年級選項
        /// </summary>
        /// <param name="items"></param>
        public void SetSelectItems(List<string> items)
        {
            cbx01.Items.Clear();
            cbx02.Items.Clear();
            cbx03.Items.Clear();
            cbx04.Items.Clear();
            cbx05.Items.Clear();
            cbx06.Items.Clear();
            if (items.Count > 0)
            {
                cbx01.Items.AddRange(items.ToArray());
                cbx02.Items.AddRange(items.ToArray());
                cbx03.Items.AddRange(items.ToArray());
                cbx04.Items.AddRange(items.ToArray());
                cbx05.Items.AddRange(items.ToArray());
                cbx06.Items.AddRange(items.ToArray());
            }
        }
              

        private void StudABCard_YearlyForm_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;
            LoadData();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            SetData();
      
            this.Close();
        }

        /// <summary>
        /// 取得目前資料
        /// </summary>
        /// <returns></returns>
        public UDTYearlyDataDef GetData()
        {
            SetData();
            return _data;
        }

        private void LoadData()
        {
            cbx01.Text = _data.G1;
            cbx02.Text = _data.G2;
            cbx03.Text = _data.G3;
            cbx04.Text = _data.G4;
            cbx05.Text = _data.G5;
            cbx06.Text = _data.G6;        
        }

        private void SetData()
        {
            _data.G1 = cbx01.Text;
            _data.G2 = cbx02.Text;
            _data.G3 = cbx03.Text;
            _data.G4 = cbx04.Text;
            _data.G5 = cbx05.Text;
            _data.G6 = cbx06.Text;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }
    }
}
